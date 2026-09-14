import streamlit as st
import pandas as pd
import datetime
import requests
from bs4 import BeautifulSoup
import holidays
import plotly.express as px
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
from openpyxl.chart import BarChart, PieChart, Reference
from openpyxl.formatting.rule import DataBarRule

st.set_page_config(page_title="SMS Schedule Agent", layout="wide")

st.title("🤖 ИИ-Агент: Генератор SMS-графика проекта")
st.write("Выберите наиболее удобное представление графика для работы и выгрузки в Excel")

@st.cache_data
def get_rf_holidays():
    url = "https://www.consultant.ru/law/ref/calendar/proizvodstvennye/"
    headers = {"User-Agent": "Mozilla/5.0"}
    holiday_dates = set()
    try:
        response = requests.get(url, headers=headers, timeout=5)
        if response.status_code == 200:
            pass
    except Exception:
        pass
        
    current_year = datetime.datetime.now().year
    ru_holidays = holidays.RU(years=[current_year - 2, current_year - 1, current_year, current_year + 1, current_year + 2, current_year + 3])
    for d in ru_holidays.keys():
        holiday_dates.add(d)
    return holiday_dates

def is_business_day(date_val, holiday_dates):
    if date_val.weekday() >= 5 or date_val in holiday_dates:
        return False
    return True

def get_next_business_day(date_val, holiday_dates):
    cur = date_val
    while not is_business_day(cur, holiday_dates):
        cur += datetime.timedelta(days=1)
    return cur

def extract_start_milestone(xls):
    sheet_name = "START PROJECT TOGF-ENG-007-02"
    if sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                cell_val = df.iloc[r, c]
                if pd.notna(cell_val) and str(cell_val).strip().upper() not in ['N/A', 'NONE', 'CLOSED']:
                    dt = pd.to_datetime(cell_val, errors='coerce')
                    if pd.notna(dt) and dt.year > 2000:
                        return dt.date()
    return None

# ==========================================
# ВАРИАНТ A: Дорожная карта (Roadmap & Milestones)
# ==========================================
def generate_excel_variant_a(schedule_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Дорожная карта"
    ws.views.sheetView[0].showGridLines = True
    
    # Шапка
    ws.merge_cells("A1:F1")
    ws["A1"] = "ДОРОЖНАЯ КАРТА И КЛЮЧЕВЫЕ ВЕХИ ПРОЕКТА"
    ws["A1"].font = Font(name="Calibri", size=14, bold=True, color="FFFFFF")
    ws["A1"].fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center")
    ws.row_dimensions[1].height = 35

    # Таблица вех по фазам
    df_sched = pd.DataFrame(schedule_data)
    phase_summary = df_sched.groupby('Фаза проекта').agg(
        Старт=('Start', 'min'),
        Финиш=('Finish', 'max'),
        Кол_во_задач=('№', 'count')
    ).reset_index()

    ws.cell(row=3, column=1, value="СВОДКА ПО ФАЗАМ (MILESTONES)").font = Font(bold=True, size=11, color="1F4E78")
    
    headers_m = ["Фаза проекта", "Дата начала", "Дата окончания", "Кол-во задач", "Статус"]
    for c_idx, h in enumerate(headers_m, 1):
        cell = ws.cell(row=4, column=c_idx, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")

    for r_idx, row in phase_summary.iterrows():
        curr_row = 5 + r_idx
        ws.cell(row=curr_row, column=1, value=row['Фаза проекта'])
        ws.cell(row=curr_row, column=2, value=row['Старт'].strftime("%d.%m.%Y")).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_row, column=3, value=row['Финиш'].strftime("%d.%m.%Y")).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_row, column=4, value=row['Кол_во_задач']).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_row, column=5, value="Запланировано").alignment = Alignment(horizontal="center")

    # Реестр всех задач
    start_task_row = 7 + len(phase_summary)
    ws.cell(row=start_task_row-1, column=1, value="ПОДРОБНЫЙ РЕЕСТР ЗАДАЧ").font = Font(bold=True, size=11, color="1F4E78")
    
    headers_t = ["№", "Фаза проекта", "Описание работы (DESCRIPTION)", "Дата выполнения", "День недели"]
    for c_idx, h in enumerate(headers_t, 1):
        cell = ws.cell(row=start_task_row, column=c_idx, value=h)
        cell.font = Font(bold=True, color="FFFFFF")
        cell.fill = PatternFill(start_color="595959", end_color="595959", fill_type="solid")
        cell.alignment = Alignment(horizontal="center")

    days_ru = ["Понедельник", "Вторник", "Среда", "Четверг", "Пятница", "Суббота", "Воскресенье"]
    for r_idx, item in enumerate(schedule_data, start_task_row+1):
        ws.cell(row=r_idx, column=1, value=item['№']).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=2, value=item['Фаза проекта'])
        ws.cell(row=r_idx, column=3, value=item['DESCRIPTION'])
        ws.cell(row=r_idx, column=4, value=item['Start'].strftime("%d.%m.%Y")).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=5, value=days_ru[item['Start'].weekday()]).alignment = Alignment(horizontal="center")

    ws.column_dimensions['A'].width = 8
    ws.column_dimensions['B'].width = 32
    ws.column_dimensions['C'].width = 55
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 16

    filename = "Roadmap_Milestones.xlsx"
    wb.save(filename)
    return filename


# ==========================================
# ВАРИАНТ B: Смарт-таблица с Индикаторами (Data Bars)
# ==========================================
def generate_excel_variant_b(schedule_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Смарт-График"
    ws.views.sheetView[0].showGridLines = True

    headers = ["№", "Фаза проекта", "Описание задачи", "Дата начала", "Дата окончания", "Рабочих дней", "Визуальный вес"]
    ws.append(headers)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    for col_idx in range(1, len(headers)+1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = Font(bold=True, color="FFFFFF")
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for item in schedule_data:
        ws.append([
            item['№'],
            item['Фаза проекта'],
            item['DESCRIPTION'],
            item['Start'].strftime("%d.%m.%Y"),
            item['Finish'].strftime("%d.%m.%Y"),
            1,
            item['№'] # Индикатор последовательности
        ])

    # Добавление цветного DataBar индикатора
    rule = DataBarRule(start_type='num', start_value=1, end_type='num', end_value=len(schedule_data),
                       color="638EC6", showValue="none", minLength=None, maxLength=None)
    ws.conditional_formatting.add(f"G2:G{len(schedule_data)+1}", rule)

    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 50
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 14
    ws.column_dimensions['F'].width = 15
    ws.column_dimensions['G'].width = 20

    filename = "Smart_Schedule.xlsx"
    wb.save(filename)
    return filename


# ==========================================
# ВАРИАНТ C: Управленческий Дашборд
# ==========================================
def generate_excel_variant_c(schedule_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Дашборд проекта"
    ws.views.sheetView[0].showGridLines = True

    # KPI Блоки
    df_s = pd.DataFrame(schedule_data)
    total_tasks = len(df_s)
    start_dt = df_s['Start'].min().strftime("%d.%m.%Y")
    end_dt = df_s['Finish'].max().strftime("%d.%m.%Y")
    total_days = (df_s['Finish'].max() - df_s['Start'].min()).days + 1

    ws.merge_cells("A1:B2")
    ws["A1"] = f"ВСЕГО ЗАДАЧ\n{total_tasks}"
    ws["A1"].font = Font(size=12, bold=True, color="1F4E78")
    ws["A1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.merge_cells("C1:D2")
    ws["C1"] = f"СТАРТ ПРОЕКТА\n{start_dt}"
    ws["C1"].font = Font(size=12, bold=True, color="2E75B6")
    ws["C1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    ws.merge_cells("E1:F2")
    ws["E1"] = f"ФИНИШ ПРОЕКТА\n{end_dt}"
    ws["E1"].font = Font(size=12, bold=True, color="C65911")
    ws["E1"].alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Таблица распределения по фазам
    ws.cell(row=4, column=1, value="Фаза проекта").font = Font(bold=True)
    ws.cell(row=4, column=2, value="Количество задач").font = Font(bold=True)
    
    phase_counts = df_s['Фаза проекта'].value_counts()
    for idx, (p_name, count) in enumerate(phase_counts.items(), 5):
        ws.cell(row=idx, column=1, value=p_name)
        ws.cell(row=idx, column=2, value=count)

    # Диаграмма распределения задач
    pie = PieChart()
    labels = Reference(ws, min_col=1, min_row=5, max_row=4+len(phase_counts))
    data = Reference(ws, min_col=2, min_row=4, max_row=4+len(phase_counts))
    pie.add_data(data, titles_from_data=True)
    pie.set_categories(labels)
    pie.title = "Распределение объема задач по фазам"
    pie.width = 14
    pie.height = 7
    ws.add_chart(pie, "D4")

    filename = "Project_Dashboard.xlsx"
    wb.save(filename)
    return filename


# Основной интерфейс приложения
uploaded_file = st.file_uploader("Загрузите Excel-файл (PLANT_MASTER_SCHEDULE P25077.xlsx)", type=["xlsx"])
target_phases = ["PMSPR TOGF-ENG-008-06 Phase2", "PMSPR TOGF-ENG-008-06 Phase 3", "PMSPR TOGF-ENG-008-06 Phase4;5"]

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    start_date = extract_start_milestone(xls)
    
    if start_date:
        st.success(f"📅 Дата старта извлечена: **{start_date.strftime('%d.%m.%Y')}**")
        
        if st.button("🚀 Сформировать варианты отчетов"):
            tasks = []
            for sheet in target_phases:
                if sheet in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet, header=None)
                    desc_col = None
                    for r in range(min(25, df.shape[0])):
                        for c in range(df.shape[1]):
                            if str(df.iloc[r, c]).strip().upper() == 'DESCRIPTION':
                                desc_col = c
                                break
                        if desc_col is not None:
                            break
                    if desc_col is not None:
                        for val in df.iloc[10:, desc_col].dropna():
                            val_str = str(val).strip()
                            if val_str and val_str.upper() != 'DESCRIPTION':
                                tasks.append({'Phase': sheet, 'DESCRIPTION': val_str})

            if tasks:
                holiday_dates = get_rf_holidays()
                current_date = start_date
                schedule_data = []
                
                for idx, row in enumerate(tasks):
                    current_date = get_next_business_day(current_date, holiday_dates)
                    schedule_data.append({
                        '№': idx + 1,
                        'Фаза проекта': row['Phase'],
                        'DESCRIPTION': row['DESCRIPTION'],
                        'Start': current_date,
                        'Finish': current_date
                    })
                    current_date = get_next_business_day(current_date + datetime.timedelta(days=1), holiday_dates)

                # Генерация трех новых типов файлов
                file_a = generate_excel_variant_a(schedule_data)
                file_b = generate_excel_variant_b(schedule_data)
                file_c = generate_excel_variant_c(schedule_data)

                tab1, tab2, tab3 = st.tabs(["📌 A. Дорожная карта", "📊 B. Смарт-реестр", "📈 C. Дашборд"])

                with tab1:
                    st.write("### Вариант A: Дорожная карта с ключевыми вехами")
                    st.write("Сворачивает график в наглядные этапы и дает удобную таблицу для печати.")
                    with open(file_a, "rb") as f:
                        st.download_button("📥 Скачать Дорожную карту (.xlsx)", f, file_name="Roadmap_Milestones.xlsx")

                with tab2:
                    st.write("### Вариант B: Смарт-реестр с цветовыми барами")
                    st.write("Чистая таблица с встроенными градиентными индикаторами длительности задач.")
                    with open(file_b, "rb") as f:
                        st.download_button("📥 Скачать Смарт-реестр (.xlsx)", f, file_name="Smart_Schedule.xlsx")

                with tab3:
                    st.write("### Вариант C: Управленческий Дашборд")
                    st.write("KPI-карточки сроков и круговая диаграмма объема работ по этапам.")
                    with open(file_c, "rb") as f:
                        st.download_button("📥 Скачать Дашборд (.xlsx)", f, file_name="Project_Dashboard.xlsx")
            else:
                st.error("Не удалось извлечь задачи из столбцов DESCRIPTION.")
