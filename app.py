# app.py

import streamlit as st
import pandas as pd
import gspread
import tempfile
import os
import time
from datetime import datetime, timedelta

from google.oauth2.service_account import Credentials

from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    HRFlowable
)

from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm

from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase import pdfmetrics

# ---------------- PDF FONT ----------------
try:
    pdfmetrics.registerFont(UnicodeCIDFont('HYSMyeongJo-Medium'))
except:
    pass

# ---------------- PAGE ----------------
st.set_page_config(page_title="🚚 Отгрузка маршрутов", layout="wide")

# ---------------- SETTINGS ----------------
SHEET_ID = "1hKZ8ggNLW-OY1bV8xAW7PKl50Fof2co86oxGK92YPAA"
SHEET_NAME = "Маршруты New"

# Инициализация session state
if 'shipment_completed' not in st.session_state:
    st.session_state.shipment_completed = False
if 'pdf_file' not in st.session_state:
    st.session_state.pdf_file = None
if 'selected_routes' not in st.session_state:
    st.session_state.selected_routes = []
if 'rollback_mode' not in st.session_state:
    st.session_state.rollback_mode = False

# ---------------- CSS ----------------
st.markdown("""
<style>
.main { background-color: #F3F4F6; }
div[data-testid="metric-container"] {
    background: white;
    border-radius: 16px;
    padding: 20px;
    box-shadow: 0 2px 12px rgba(0,0,0,0.05);
}
.stButton>button {
    background-color: #16A34A;
    color: white;
    border-radius: 12px;
    border: none;
    padding: 14px 22px;
    font-weight: 700;
    width: 100%;
    font-size: 16px;
}
.stButton>button:hover { background-color: #15803D; }
div[data-testid="stDownloadButton"] > button {
    background-color: #3B82F6;
    color: white;
    border-radius: 12px;
    border: none;
    padding: 14px 22px;
    font-weight: 700;
    width: 100%;
    font-size: 16px;
}
.stButton > button[kind="secondary"] {
    background-color: #EF4444;
}
</style>
""", unsafe_allow_html=True)

# ---------------- GOOGLE SHEETS ----------------
def connect_sheet():
    try:
        if "gcp_service_account" not in st.secrets:
            st.error("❌ gcp_service_account не найден в secrets!")
            return None
        
        creds_dict = dict(st.secrets["gcp_service_account"])
        if 'private_key' in creds_dict:
            creds_dict['private_key'] = creds_dict['private_key'].replace('\\n', '\n')
        
        scope = ["https://www.googleapis.com/auth/spreadsheets", "https://www.googleapis.com/auth/drive"]
        credentials = Credentials.from_service_account_info(creds_dict, scopes=scope)
        client = gspread.authorize(credentials)
        sheet = client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)
        return sheet
    except Exception as e:
        st.error(f"❌ Ошибка подключения: {str(e)}")
        return None

def get_data():
    sheet = connect_sheet()
    if sheet is None:
        return pd.DataFrame()
    try:
        data = sheet.get_all_records()
        return pd.DataFrame(data)
    except Exception as e:
        st.error(f"❌ Ошибка чтения: {str(e)}")
        return pd.DataFrame()

def update_routes_batch(routes_list, car_number, driver, plomb, trip_number):
    """Пакетное обновление данных - БЕЗ ПРЕВЫШЕНИЯ ЛИМИТОВ"""
    sheet = connect_sheet()
    if sheet is None:
        return False
    
    try:
        # Получаем все данные
        all_data = sheet.get_all_values()
        headers = all_data[0]
        
        # Находим индексы колонок
        col_indices = {
            'status': headers.index("Статус отгрузки"),
            'fact': headers.index("Дата отгрузки факт") if "Дата отгрузки факт" in headers else None,
            'car': headers.index("Номер машины") if "Номер машины" in headers else None,
            'driver': headers.index("Водитель") if "Водитель" in headers else None,
            'plomb': headers.index("№ пломбы") if "№ пломбы" in headers else None,
            'trip': headers.index("Рейс") if "Рейс" in headers else None,
        }
        
        # Добавляем недостающие колонки
        if col_indices['fact'] is None:
            sheet.add_cols(5)
            all_data = sheet.get_all_values()
            headers = all_data[0]
            col_indices['fact'] = headers.index("Дата отгрузки факт")
            col_indices['car'] = headers.index("Номер машины")
            col_indices['driver'] = headers.index("Водитель")
            col_indices['plomb'] = headers.index("№ пломбы")
            col_indices['trip'] = headers.index("Рейс")
        
        # UTC+5
        now_utc_plus_5 = datetime.now() + timedelta(hours=5)
        fact_date = now_utc_plus_5.strftime("%d.%m.%Y %H:%M")
        
        # Подготавливаем обновления
        updates = []
        route_indices = [str(r) for r in routes_list]
        
        for idx, row in enumerate(all_data[1:], start=2):
            if len(row) > headers.index("Номер маршрута") and str(row[headers.index("Номер маршрута")]) in route_indices:
                updates.append({
                    'range': f'{chr(65 + col_indices["status"])}{idx}',
                    'value': "ОТГРУЖЕН"
                })
                updates.append({
                    'range': f'{chr(65 + col_indices["fact"])}{idx}',
                    'value': fact_date
                })
                updates.append({
                    'range': f'{chr(65 + col_indices["car"])}{idx}',
                    'value': car_number
                })
                updates.append({
                    'range': f'{chr(65 + col_indices["driver"])}{idx}',
                    'value': driver
                })
                updates.append({
                    'range': f'{chr(65 + col_indices["plomb"])}{idx}',
                    'value': plomb
                })
                updates.append({
                    'range': f'{chr(65 + col_indices["trip"])}{idx}',
                    'value': trip_number
                })
        
        # Пакетное обновление (максимум 100 ячеек за раз)
        if updates:
            batch = []
            for i, update in enumerate(updates):
                batch.append(update)
                # Отправляем каждые 50 обновлений или в конце
                if len(batch) >= 50 or i == len(updates) - 1:
                    sheet.batch_update([{
                        'range': u['range'],
                        'values': [[u['value']]]
                    } for u in batch])
                    batch = []
                    time.sleep(0.5)  # Пауза, чтобы не превысить лимит
        
        return True
        
    except Exception as e:
        st.error(f"❌ Ошибка обновления: {str(e)}")
        return False

def rollback_orders_batch(order_numbers):
    """Пакетный откат заказов"""
    sheet = connect_sheet()
    if sheet is None:
        return False
    
    try:
        all_data = sheet.get_all_values()
        headers = all_data[0]
        
        col_indices = {
            'status': headers.index("Статус отгрузки"),
            'fact': headers.index("Дата отгрузки факт"),
            'car': headers.index("Номер машины"),
            'driver': headers.index("Водитель"),
            'plomb': headers.index("№ пломбы"),
            'trip': headers.index("Рейс"),
        }
        
        updates = []
        order_strs = [str(o) for o in order_numbers]
        
        for idx, row in enumerate(all_data[1:], start=2):
            if len(row) > headers.index("№ заказа") and str(row[headers.index("№ заказа")]) in order_strs:
                updates.append({'range': f'{chr(65 + col_indices["status"])}{idx}', 'value': ""})
                updates.append({'range': f'{chr(65 + col_indices["fact"])}{idx}', 'value': ""})
                updates.append({'range': f'{chr(65 + col_indices["car"])}{idx}', 'value': ""})
                updates.append({'range': f'{chr(65 + col_indices["driver"])}{idx}', 'value': ""})
                updates.append({'range': f'{chr(65 + col_indices["plomb"])}{idx}', 'value': ""})
                updates.append({'range': f'{chr(65 + col_indices["trip"])}{idx}', 'value': ""})
        
        if updates:
            batch = []
            for i, update in enumerate(updates):
                batch.append(update)
                if len(batch) >= 50 or i == len(updates) - 1:
                    sheet.batch_update([{'range': u['range'], 'values': [[u['value']]]} for u in batch])
                    batch = []
                    time.sleep(0.5)
        
        return True
        
    except Exception as e:
        st.error(f"❌ Ошибка отката: {str(e)}")
        return False

# ---------------- PDF GENERATION ----------------
def generate_delivery_pdf(all_data, routes_list, driver, car, plomb, trip_number):
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
        filename = tmp_file.name
    
    doc = SimpleDocTemplate(
        filename,
        pagesize=landscape(A4),
        leftMargin=10*mm,
        rightMargin=10*mm,
        topMargin=15*mm,
        bottomMargin=15*mm
    )

    styles = getSampleStyleSheet()
    
    styleTitle = ParagraphStyle('CustomTitle', parent=styles['Normal'], fontName='HYSMyeongJo-Medium', fontSize=14, leading=18, alignment=1, spaceAfter=10, textColor=colors.black)
    styleInfo = ParagraphStyle('CustomInfo', parent=styles['Normal'], fontName='HYSMyeongJo-Medium', fontSize=10, leading=14, spaceAfter=3, textColor=colors.black)
    styleCell = ParagraphStyle('CustomCell', parent=styles['Normal'], fontName='HYSMyeongJo-Medium', fontSize=9, leading=11, textColor=colors.black)
    styleBold = ParagraphStyle('CustomBold', parent=styles['Normal'], fontName='HYSMyeongJo-Medium', fontSize=9, leading=11, textColor=colors.black)
    styleHeader = ParagraphStyle('CustomHeader', parent=styles['Normal'], fontName='HYSMyeongJo-Medium', fontSize=9, leading=12, alignment=1, textColor=colors.black)

    elements = []
    total_points = len(all_data)

    elements.append(Paragraph("МАРШРУТНЫЙ ЛИСТ ДОСТАВКИ", styleTitle))
    elements.append(Spacer(1, 4))
    
    elements.append(Paragraph(f"<b>Рейс:</b> {trip_number}", styleInfo))
    elements.append(Paragraph(f"<b>Водитель:</b> {driver}", styleInfo))
    elements.append(Paragraph(f"<b>Машина:</b> {car}", styleInfo))
    elements.append(Paragraph(f"<b>Пломба:</b> {plomb}", styleInfo))
    elements.append(Paragraph(f"<b>Заказов:</b> {total_points}", styleInfo))
    
    elements.append(Spacer(1, 6))
    elements.append(HRFlowable(width="100%", thickness=1.2, color=colors.black))
    elements.append(Spacer(1, 6))

    headers = ["№ заказа", "Магазин", "Адрес", "Пломба", "Выдано коробок", "Получено коробок", "Подпись, печать, комментарии", "Подпись водителя"]
    table_data = [[Paragraph(h, styleHeader) for h in headers]]

    for _, row in all_data.iterrows():
        order_num = f"<b>{row['№ заказа']}</b>"
        shop_name = str(row["Название магазина"])[:60]
        address = str(row["Адрес магазина"])[:80]
        
        table_data.append([
            Paragraph(order_num, styleBold),
            Paragraph(shop_name, styleCell),
            Paragraph(address, styleCell),
            Paragraph(" ", styleCell),
            Paragraph(" ", styleCell),
            Paragraph(" ", styleCell),
            Paragraph(" ", styleCell),
            Paragraph(" ", styleCell)
        ])

    table = Table(table_data, colWidths=[20*mm, 50*mm, 70*mm, 20*mm, 25*mm, 25*mm, 45*mm, 30*mm], repeatRows=1)
    
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E0E0E0")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
        ('FONTNAME', (0,0), (-1,0), 'HYSMyeongJo-Medium'),
        ('FONTSIZE', (0,0), (-1,0), 9),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('VALIGN', (0,0), (-1,0), 'MIDDLE'),
        ('FONTNAME', (0,1), (-1,-1), 'HYSMyeongJo-Medium'),
        ('FONTSIZE', (0,1), (-1,-1), 9),
        ('VALIGN', (0,1), (-1,-1), 'MIDDLE'),
        ('GRID', (0,0), (-1,-1), 1, colors.black),
        ('BOX', (0,0), (-1,-1), 1.5, colors.black),
        ('TOPPADDING', (0,0), (-1,-1), 5),
        ('BOTTOMPADDING', (0,0), (-1,-1), 5),
        ('LEFTPADDING', (0,0), (-1,-1), 4),
        ('RIGHTPADDING', (0,0), (-1,-1), 4),
        ('ALIGN', (0,1), (0,-1), 'CENTER'),
        ('ALIGN', (1,1), (2,-1), 'LEFT'),
        ('ALIGN', (3,1), (-1,-1), 'CENTER'),
    ]))

    elements.append(table)
    elements.append(Spacer(1, 20))
    
    styleSignatures = ParagraphStyle('CustomSignatures', parent=styles['Normal'], fontName='HYSMyeongJo-Medium', fontSize=10, leading=14, textColor=colors.black)
    elements.append(Spacer(1, 10))
    elements.append(Paragraph("<b>Подпись водителя:</b> _________________________ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <b>Подпись ответственного:</b> _________________________", styleSignatures))

    doc.build(elements)
    return filename

# ---------------- MAIN ----------------
st.title("🚚 Система отгрузки маршрутов")

with st.spinner("Загрузка данных..."):
    df = get_data()

if df.empty:
    st.error("❌ Нет данных")
    st.stop()

not_shipped = df[df["Статус отгрузки"] != "ОТГРУЖЕН"]
shipped = df[df["Статус отгрузки"] == "ОТГРУЖЕН"]

col1, col2, col3, col4 = st.columns(4)
col1.metric("Не отгружено", not_shipped["Номер маршрута"].nunique())
col2.metric("Отгружено", shipped["Номер маршрута"].nunique())
col3.metric("Точек", len(not_shipped))
col4.metric("Кол-во шт", int(not_shipped["кол-во штук в заказе"].sum()) if not not_shipped.empty else 0)

st.divider()

if st.button("🔄 Откатить заказы", type="secondary"):
    st.session_state.rollback_mode = True
    st.rerun()

if st.session_state.get('rollback_mode', False):
    st.subheader("🔄 Режим отката")
    shipped_orders = shipped[shipped["Статус отгрузки"] == "ОТГРУЖЕН"]
    
    if len(shipped_orders) > 0:
        orders_to_rollback = st.multiselect("Выберите заказы", options=shipped_orders["№ заказа"].tolist())
        
        col1, col2 = st.columns(2)
        if col1.button("✅ Подтвердить"):
            if orders_to_rollback and rollback_orders_batch(orders_to_rollback):
                st.success(f"Откачено {len(orders_to_rollback)} заказов")
                st.session_state.rollback_mode = False
                st.rerun()
        if col2.button("❌ Отмена"):
            st.session_state.rollback_mode = False
            st.rerun()
    else:
        st.info("Нет отгруженных заказов")
        if st.button("Назад"):
            st.session_state.rollback_mode = False
            st.rerun()

elif not st.session_state.get('shipment_completed', False):
    col1, col2 = st.columns(2)
    
    with col1:
        st.subheader("🚛 Данные")
        car_number = st.text_input("Номер машины")
        driver = st.text_input("Водитель")
        plomb = st.text_input("№ пломбы")
        trip_number = st.text_input("Рейс")
    
    with col2:
        st.subheader("📦 Маршруты")
        if not not_shipped.empty:
            routes = sorted(not_shipped["Номер маршрута"].dropna().unique())
            selected_routes = st.multiselect("Выберите маршруты", routes)
        else:
            selected_routes = []
    
    if selected_routes:
        st.subheader("📋 Детали")
        details = not_shipped[not_shipped["Номер маршрута"].isin(selected_routes)]
        st.dataframe(details[["№ заказа", "Название магазина", "Адрес магазина", "Номер маршрута"]], use_container_width=True)
    
    if st.button("✅ ОТГРУЗИТЬ", type="primary"):
        if not all([car_number, driver, plomb, trip_number, selected_routes]):
            st.warning("Заполните все поля")
        else:
            with st.spinner("Создание PDF..."):
                data_for_pdf = not_shipped[not_shipped["Номер маршрута"].isin(selected_routes)]
                if update_routes_batch(selected_routes, car_number, driver, plomb, trip_number):
                    pdf_file = generate_delivery_pdf(data_for_pdf, selected_routes, driver, car_number, plomb, trip_number)
                    st.session_state.pdf_file = pdf_file
                    st.session_state.shipment_completed = True
                    st.session_state.selected_routes = selected_routes
                    st.rerun()

elif st.session_state.get('shipment_completed') and st.session_state.get('pdf_file'):
    st.success(f"✅ Отгружено: {', '.join(str(r) for r in st.session_state.selected_routes)}")
    
    if os.path.exists(st.session_state.pdf_file):
        with open(st.session_state.pdf_file, "rb") as f:
            st.download_button("📄 Скачать PDF", f, file_name=f"route_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf", mime="application/pdf")
    
    if st.button("🔄 Новая отгрузка"):
        try:
            if st.session_state.pdf_file and os.path.exists(st.session_state.pdf_file):
                os.unlink(st.session_state.pdf_file)
        except:
            pass
        st.session_state.shipment_completed = False
        st.session_state.pdf_file = None
        st.session_state.selected_routes = []
        st.rerun()
