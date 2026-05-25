# app.py

import streamlit as st
import pandas as pd
import gspread
import tempfile
import os
from datetime import datetime, timedelta

from oauth2client.service_account import ServiceAccountCredentials

from reportlab.platypus import (
    SimpleDocTemplate,
    Table,
    TableStyle,
    Paragraph,
    Spacer,
    HRFlowable,
    PageBreak
)

from reportlab.lib import colors
from reportlab.lib.styles import getSampleStyleSheet, ParagraphStyle
from reportlab.lib.pagesizes import A4, landscape
from reportlab.lib.units import mm, cm

from reportlab.pdfbase.cidfonts import UnicodeCIDFont
from reportlab.pdfbase import pdfmetrics

# ---------------- PDF FONT ----------------

pdfmetrics.registerFont(
    UnicodeCIDFont('HYSMyeongJo-Medium')
)

# ---------------- PAGE ----------------

st.set_page_config(
    page_title="🚚 Отгрузка маршрутов",
    layout="wide"
)

# ---------------- SETTINGS ----------------

SHEET_ID = "1hKZ8ggNLW-OY1bV8xAW7PKl50Fof2co86oxGK92YPAA"
SHEET_NAME = "Маршруты New"  # Изменено на "Маршруты New"

# Инициализация session state
if 'shipment_completed' not in st.session_state:
    st.session_state.shipment_completed = False

if 'pdf_file' not in st.session_state:
    st.session_state.pdf_file = None

if 'selected_routes' not in st.session_state:
    st.session_state.selected_routes = []

if 'rollback_mode' not in st.session_state:
    st.session_state.rollback_mode = False

if 'rolled_back_orders' not in st.session_state:
    st.session_state.rolled_back_orders = []

# ---------------- CSS ----------------

st.markdown("""
<style>
.main {
    background-color: #F3F4F6;
}

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

.stButton>button:hover {
    background-color: #15803D;
}

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

div[data-testid="stDownloadButton"] > button:hover {
    background-color: #2563EB;
}

.stButton > button[kind="secondary"] {
    background-color: #EF4444;
}

.stButton > button[kind="secondary"]:hover {
    background-color: #DC2626;
}
</style>
""", unsafe_allow_html=True)

# ---------------- GOOGLE SHEETS ----------------

def connect_sheet():
    scope = [
        "https://spreadsheets.google.com/feeds",
        "https://www.googleapis.com/auth/drive"
    ]

    creds = ServiceAccountCredentials.from_json_keyfile_dict(
        st.secrets["gcp_service_account"],
        scope
    )

    client = gspread.authorize(creds)
    sheet = client.open_by_key(SHEET_ID).worksheet(SHEET_NAME)
    return sheet

def get_data():
    sheet = connect_sheet()
    data = sheet.get_all_records()
    return pd.DataFrame(data)

def update_routes(routes_list, car_number, driver, plomb, trip_number):
    sheet = connect_sheet()
    data = sheet.get_all_records()
    headers = sheet.row_values(1)

    extra_columns = ["Номер машины", "Водитель", "№ пломбы", "Дата отгрузки факт", "Рейс"]
    current_headers = headers.copy()

    for col_name in extra_columns:
        if col_name not in current_headers:
            sheet.update_cell(1, len(current_headers) + 1, col_name)
            current_headers.append(col_name)

    headers = sheet.row_values(1)
    status_col = headers.index("Статус отгрузки") + 1
    fact_col = headers.index("Дата отгрузки факт") + 1
    car_col = headers.index("Номер машины") + 1
    driver_col = headers.index("Водитель") + 1
    plomb_col = headers.index("№ пломбы") + 1
    trip_col = headers.index("Рейс") + 1

    # Исправляем часовой пояс (UTC+5) - добавляем 5 часов
    now_utc_plus_5 = datetime.now() + timedelta(hours=5)
    
    for idx, row in enumerate(data, start=2):
        if row["Номер маршрута"] in routes_list:
            sheet.update_cell(idx, status_col, "ОТГРУЖЕН")
            sheet.update_cell(idx, fact_col, now_utc_plus_5.strftime("%d.%m.%Y %H:%M"))
            sheet.update_cell(idx, car_col, car_number)
            sheet.update_cell(idx, driver_col, driver)
            sheet.update_cell(idx, plomb_col, plomb)
            sheet.update_cell(idx, trip_col, trip_number)

def rollback_orders(order_numbers):
    """Функция для отката заказов"""
    sheet = connect_sheet()
    data = sheet.get_all_records()
    headers = sheet.row_values(1)
    
    status_col = headers.index("Статус отгрузки") + 1
    fact_col = headers.index("Дата отгрузки факт") + 1
    car_col = headers.index("Номер машины") + 1
    driver_col = headers.index("Водитель") + 1
    plomb_col = headers.index("№ пломбы") + 1
    trip_col = headers.index("Рейс") + 1
    
    for idx, row in enumerate(data, start=2):
        if row["№ заказа"] in order_numbers:
            sheet.update_cell(idx, status_col, "")  # Очищаем статус
            sheet.update_cell(idx, fact_col, "")   # Очищаем дату
            sheet.update_cell(idx, car_col, "")    # Очищаем машину
            sheet.update_cell(idx, driver_col, "") # Очищаем водителя
            sheet.update_cell(idx, plomb_col, "")  # Очищаем пломбу
            sheet.update_cell(idx, trip_col, "")   # Очищаем рейс
    return True

def get_shipped_routes():
    """Получить список отгруженных маршрутов"""
    df = get_data()
    shipped = df[df["Статус отгрузки"] == "ОТГРУЖЕН"]
    return shipped

# ---------------- PDF GENERATION ----------------

def generate_delivery_pdf(all_data, routes_list, driver, car, plomb, trip_number):
    """Генерация ОДНОГО PDF для всех выбранных маршрутов"""
    
    with tempfile.NamedTemporaryFile(delete=False, suffix='.pdf') as tmp_file:
        filename = tmp_file.name
    
    # Создаем документ с минимальными отступами для максимальной растяжки
    doc = SimpleDocTemplate(
        filename,
        pagesize=landscape(A4),
        leftMargin=2*mm,
        rightMargin=2*mm,
        topMargin=10*mm,
        bottomMargin=15*mm
    )

    styles = getSampleStyleSheet()
    
    # Стили для кириллицы
    styleTitle = ParagraphStyle(
        'CustomTitle',
        parent=styles['Normal'],
        fontName='HYSMyeongJo-Medium',
        fontSize=14,
        leading=18,
        alignment=1,
        spaceAfter=10,
        spaceBefore=4,
        textColor=colors.black
    )
    
    styleInfo = ParagraphStyle(
        'CustomInfo',
        parent=styles['Normal'],
        fontName='HYSMyeongJo-Medium',
        fontSize=9,
        leading=12,
        spaceAfter=3,
        textColor=colors.black
    )
    
    styleCell = ParagraphStyle(
        'CustomCell',
        parent=styles['Normal'],
        fontName='HYSMyeongJo-Medium',
        fontSize=8,
        leading=10,
        textColor=colors.black
    )
    
    styleBold = ParagraphStyle(
        'CustomBold',
        parent=styles['Normal'],
        fontName='HYSMyeongJo-Medium',
        fontSize=9,
        leading=11,
        alignment=0,
        textColor=colors.black
    )
    
    styleHeader = ParagraphStyle(
        'CustomHeader',
        parent=styles['Normal'],
        fontName='HYSMyeongJo-Medium',
        fontSize=8,
        leading=10,
        alignment=1,
        textColor=colors.black
    )

    elements = []
    total_points = len(all_data)
    routes_text = ", ".join(str(r) for r in routes_list)

    # ЗАГОЛОВОК - без маршрута и даты
    title = Paragraph("МАРШРУТНЫЙ ЛИСТ ДОСТАВКИ", styleTitle)
    elements.append(title)
    elements.append(Spacer(1, 4))
    
    # ИНФОРМАЦИЯ - убрали маршрут и дату
    info1 = Paragraph(f"<b>Рейс:</b> {trip_number}", styleInfo)
    elements.append(info1)
    
    info2 = Paragraph(f"<b>Водитель:</b> {driver}", styleInfo)
    elements.append(info2)
    
    info3 = Paragraph(f"<b>Машина:</b> {car}", styleInfo)
    elements.append(info3)
    
    info4 = Paragraph(f"<b>Пломба:</b> {plomb}", styleInfo)
    elements.append(info4)
    
    info5 = Paragraph(f"<b>Заказов:</b> {total_points}", styleInfo)
    elements.append(info5)
    
    elements.append(Spacer(1, 6))
    elements.append(HRFlowable(width="100%", thickness=1.2, color=colors.black))
    elements.append(Spacer(1, 6))

    # ТАБЛИЦА - заголовки с переносами
    headers = [
        "№\nзаказа",
        "Магазин",
        "Адрес",
        "Пломба",
        "Выдано\nкоробок",
        "Получено\nкоробок",
        "Подпись,\nпечать,\nкомментарии",
        "Подпись\nводителя"
    ]
    
    table_data = [[Paragraph(h, styleHeader) for h in headers]]

    # Данные - все ячейки пустые для заполнения
    for _, row in all_data.iterrows():
        order_num = f"<b>{row['№ заказа']}</b>"
        shop_name = str(row["Название магазина"])[:80]  # Увеличил лимит
        address = str(row["Адрес магазина"])[:120]      # Увеличил лимит
        
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

    # Максимально растянутые ширины колонок
    table = Table(
        table_data,
        colWidths=[
            18*mm,  # № заказа (увеличено)
            55*mm,  # Магазин (еще растянут)
            75*mm,  # Адрес (максимально растянут)
            20*mm,  # Пломба (увеличено)
            22*mm,  # Выдано коробок (увеличено)
            22*mm,  # Получено коробок (увеличено)
            50*mm,  # Подпись, печать, комментарии (еще растянут)
            35*mm   # Подпись водителя (еще растянут)
        ],
        repeatRows=1
    )

    # Стиль таблицы
    table.setStyle(TableStyle([
        ('BACKGROUND', (0,0), (-1,0), colors.HexColor("#E0E0E0")),
        ('TEXTCOLOR', (0,0), (-1,0), colors.black),
        ('FONTNAME', (0,0), (-1,0), 'HYSMyeongJo-Medium'),
        ('FONTSIZE', (0,0), (-1,0), 8),
        ('ALIGN', (0,0), (-1,0), 'CENTER'),
        ('VALIGN', (0,0), (-1,0), 'MIDDLE'),
        
        ('FONTNAME', (0,1), (-1,-1), 'HYSMyeongJo-Medium'),
        ('FONTSIZE', (0,1), (-1,-1), 9),
        ('VALIGN', (0,1), (-1,-1), 'MIDDLE'),
        
        ('GRID', (0,0), (-1,-1), 1.2, colors.black),
        ('BOX', (0,0), (-1,-1), 2.0, colors.black),
        
        ('TOPPADDING', (0,0), (-1,-1), 6),
        ('BOTTOMPADDING', (0,0), (-1,-1), 6),
        ('LEFTPADDING', (0,0), (-1,-1), 5),
        ('RIGHTPADDING', (0,0), (-1,-1), 5),
        
        ('ALIGN', (0,1), (0,-1), 'CENTER'),
        ('ALIGN', (1,1), (2,-1), 'LEFT'),
        ('ALIGN', (3,1), (-1,-1), 'CENTER'),
    ]))

    elements.append(table)
    elements.append(Spacer(1, 25))
    
    # ПОДПИСИ
    styleSignatures = ParagraphStyle(
        'CustomSignatures',
        parent=styles['Normal'],
        fontName='HYSMyeongJo-Medium',
        fontSize=9,
        leading=14,
        textColor=colors.black
    )
    
    elements.append(Spacer(1, 15))
    
    signatures = Paragraph(
        "<b>Подпись водителя:</b> _________________________ &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp; <b>Подпись ответственного:</b> _________________________",
        styleSignatures
    )
    elements.append(signatures)
    elements.append(Spacer(1, 10))

    doc.build(elements)
    return filename

# ---------------- MAIN ----------------

# Загрузка данных
try:
    df = get_data()
except Exception as e:
    st.error("Ошибка подключения Google Sheets")
    st.exception(e)
    st.stop()

# Фильтрация
not_shipped = df[df["Статус отгрузки"] != "ОТГРУЖЕН"]
shipped = df[df["Статус отгрузки"] == "ОТГРУЖЕН"]

# Заголовок
st.title("🚚 Система отгрузки маршрутов")

# Метрики
col1, col2, col3, col4 = st.columns(4)
col1.metric("Не отгружено", not_shipped["Номер маршрута"].nunique())
col2.metric("Отгружено", shipped["Номер маршрута"].nunique())
col3.metric("Точек", len(not_shipped))
col4.metric("Кол-во шт", int(not_shipped["кол-во штук в заказе"].sum()) if not not_shipped.empty else 0)

st.divider()

# Итоги по неотгруженным точкам
st.subheader("📊 Итоги по неотгруженным точкам")
if not not_shipped.empty:
    summary_not_shipped = not_shipped.groupby(["Название магазина", "Адрес магазина"])["кол-во штук в заказе"].sum().reset_index()
    summary_not_shipped.columns = ["Магазин", "Адрес", "Кол-во шт"]

    total_shirts = summary_not_shipped["Кол-во шт"].sum()
    summary_with_total = pd.concat([
        summary_not_shipped,
        pd.DataFrame([["ИТОГО:", "", total_shirts]], columns=summary_not_shipped.columns)
    ])

    st.dataframe(summary_with_total, use_container_width=True, height=400, hide_index=True)
else:
    st.info("Нет неотгруженных заказов")
    
st.divider()

# Режим отката
if st.button("🔄 Откатить заказы (перепечатать PDF)", type="secondary"):
    st.session_state.rollback_mode = True
    st.session_state.shipment_completed = False
    st.rerun()

if st.session_state.rollback_mode:
    st.subheader("🔄 Режим отката заказов")
    st.warning("Выберите заказы, которые нужно вернуть в статус 'Не отгружены' для перепечатки PDF")
    
    shipped_orders = shipped[shipped["Статус отгрузки"] == "ОТГРУЖЕН"]
    
    if len(shipped_orders) > 0:
        # Выбор заказов для отката
        orders_to_rollback = st.multiselect(
            "Выберите заказы для отката",
            options=shipped_orders["№ заказа"].tolist(),
            format_func=lambda x: f"Заказ {x} - Маршрут {shipped_orders[shipped_orders['№ заказа']==x]['Номер маршрута'].values[0]}"
        )
        
        col1, col2 = st.columns(2)
        with col1:
            if st.button("✅ Подтвердить откат"):
                if orders_to_rollback:
                    if rollback_orders(orders_to_rollback):
                        st.success(f"✅ {len(orders_to_rollback)} заказов откачено!")
                        st.session_state.rollback_mode = False
                        st.rerun()
                    else:
                        st.error("Ошибка при откате")
                else:
                    st.warning("Выберите заказы для отката")
        
        with col2:
            if st.button("❌ Отмена"):
                st.session_state.rollback_mode = False
                st.rerun()
    else:
        st.info("Нет отгруженных заказов для отката")
        if st.button("Назад"):
            st.session_state.rollback_mode = False
            st.rerun()

# Основная форма отгрузки
elif not st.session_state.shipment_completed:
    left, right = st.columns([1, 1])

    with left:
        st.subheader("🚛 Данные машины")
        car_number = st.text_input("Номер машины")
        driver = st.text_input("Водитель")
        plomb = st.text_input("№ пломбы")
        trip_number = st.text_input("Рейс (например: 1, 2, 3...)")

    with right:
        st.subheader("📦 Маршруты")
        if not not_shipped.empty:
            routes = sorted(not_shipped["Номер маршрута"].dropna().unique())
            selected_routes = st.multiselect("Выберите маршруты", routes)
        else:
            st.info("Нет доступных маршрутов для отгрузки")
            selected_routes = []

    # Детали выбранных маршрутов
    if selected_routes:
        st.subheader("📋 Детали маршрутов")
        details_df = not_shipped[not_shipped["Номер маршрута"].isin(selected_routes)]
        if not details_df.empty:
            st.dataframe(
                details_df[["№ заказа", "Название магазина", "Адрес магазина", "Номер маршрута", "кол-во штук в заказе"]],
                use_container_width=True,
                height=400
            )

    # Кнопка отгрузки
    if st.button("✅ ОТГРУЗИТЬ И СОЗДАТЬ PDF"):
        if not car_number:
            st.warning("Введите номер машины")
            st.stop()
        if not driver:
            st.warning("Введите ФИО водителя")
            st.stop()
        if not plomb:
            st.warning("Введите номер пломбы")
            st.stop()
        if not trip_number:
            st.warning("Введите номер рейса")
            st.stop()
        if not selected_routes:
            st.warning("Выберите маршрут")
            st.stop()

        # Собираем все данные по выбранным маршрутам
        all_selected_data = not_shipped[not_shipped["Номер маршрута"].isin(selected_routes)]
        
        # Обновляем статусы в Google Sheets
        update_routes(selected_routes, car_number, driver, plomb, trip_number)
        
        # Генерируем ОДИН PDF для всех маршрутов
        pdf_file = generate_delivery_pdf(all_selected_data, selected_routes, driver, car_number, plomb, trip_number)
        
        st.session_state.pdf_file = pdf_file
        st.session_state.shipment_completed = True
        st.session_state.selected_routes = selected_routes
        st.rerun()

# Скачивание PDF
elif st.session_state.shipment_completed and st.session_state.pdf_file:
    st.success(f"✅ Маршруты {', '.join(str(r) for r in st.session_state.selected_routes)} успешно отгружены!")
    
    # Проверяем существует ли файл
    if os.path.exists(st.session_state.pdf_file):
        with open(st.session_state.pdf_file, "rb") as f:
            st.download_button(
                label="📄 Скачать маршрутный лист (PDF)",
                data=f,
                file_name=f"Маршрутный_лист_{datetime.now().strftime('%Y%m%d_%H%M')}.pdf",
                mime="application/pdf"
            )
    else:
        st.error("Файл PDF не найден. Пожалуйста, повторите отгрузку.")
        if st.button("🔄 Повторить отгрузку"):
            st.session_state.shipment_completed = False
            st.session_state.pdf_file = None
            st.rerun()
    
    st.divider()
    
    if st.button("🔄 Начать новую отгрузку"):
        try:
            if st.session_state.pdf_file and os.path.exists(st.session_state.pdf_file):
                os.unlink(st.session_state.pdf_file)
        except:
            pass
        st.session_state.shipment_completed = False
        st.session_state.pdf_file = None
        st.session_state.selected_routes = []
        st.rerun()
else:
    # Если состояние не соответствует, сбрасываем
    st.session_state.shipment_completed = False
    st.session_state.pdf_file = None
    st.rerun()
