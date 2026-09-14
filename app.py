import streamlit as st
import pandas as pd
import datetime
import holidays
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

st.set_page_config(page_title="SMS Schedule Agent", layout="wide")

st.title("🤖 ИИ-Агент: Генератор настоящей Диаграммы Ганта в Excel")
st.write("Автоматическое построение графической диаграммы Ганта на отдельном листе Excel")

@st.cache_data
def get_rf_holidays():
    current_year = datetime.datetime.now().year
    ru_holidays = holidays.RU(years=[current_year - 2, current_year - 1, current_year, current_year + 1, current_year + 2, current_year + 3])
    return set(ru_holidays.keys())

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

def create_true_gantt_chart(schedule_data):
    wb = openpyxl.Workbook()
    
    # --- ЛИСТ 1: ТАБЛИЦА С ДАННЫМИ ---
    ws_data = wb.active
    ws_data.title = "Реестр задач"
    ws_data.views.sheetView[0].showGridLines = True

    headers = ["№", "Фаза проекта", "Описание работы (DESCRIPTION)", "Дата начала", "Дата финиша", "Смещение (дней)", "Длительность (дней)"]
    ws_data.append(headers)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    for col_idx in range(1, len(headers) + 1):
        cell = ws_data.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    base_date = schedule_data[0]['Start']

    for item in schedule_data:
        offset_days = (item['Start'] - base_date).days
        duration_days = max(1, (item['Finish'] - item['Start']).days + 1)
        
        ws_data.append([
            item['№'],
            item['Фаза проекта'],
            f"{item['№']}. {item['DESCRIPTION']}",
            item['Start'].strftime("%d.%m.%Y"),
            item['Finish'].strftime("%d.%m.%Y"),
            offset_days,
            duration_days
        ])

    ws_data.column_dimensions['A'].width = 6
    ws_data.column_dimensions['B'].width = 28
    ws_data.column_dimensions['C'].width = 50
    ws_data.column_dimensions['D'].width = 14
    ws_data.column_dimensions['E'].width = 14
    ws_data.column_dimensions['F'].width = 16
    ws_data.column_dimensions['G'].width = 18

    # --- ЛИСТ 2: НАСТОЯЩИЙ ГРАФИК (ДИАГРАММА ГАНТА) ---
    ws_chart = wb.create_sheet(title="Диаграмма Ганта")
    ws_chart.views.sheetView[0].showGridLines = True

    chart = BarChart()
    chart.type = "bar"
    chart.dir = "bar"
    chart.style = 13
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = f"График выполнения работ (Старт проекта: {base_date.strftime('%d.%m.%Y')})"
    chart.height = max(12, len(schedule_data) * 0.7)
    chart.width = 22

    # Ссылки на данные с Листа 1
    data = Reference(ws_data, min_col=6, min_row=1, max_col=7, max_row=len(schedule_data) + 1)
    cats = Reference(ws_data, min_col=3, min_row=2, max_row=len(schedule_data) + 1)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    # Делаем первому ряду (Смещение) прозрачный фон для эффекта Ганта
    if len(chart.series) > 0:
        chart.series[0].graphicalProperties.solidFill = "FFFFFF"
        chart.series[0].graphicalProperties.line.solidFill = "FFFFFF"

    # Размещаем диаграмму на весь лист
    ws_chart.add_chart(chart, "B2")

    filename = "SMS_Project_Gantt_Chart.xlsx"
    wb.save(filename)
    return filename

# --- ИНТЕРФЕЙС STREAMLIT ---
uploaded_file = st.file_uploader("Загрузите Excel-файл проекта (PLANT_MASTER_SCHEDULE P25077.xlsx)", type=["xlsx"])
target_phases = ["PMSPR TOGF-ENG-008-06 Phase2", "PMSPR TOGF-ENG-008-06 Phase 3", "PMSPR TOGF-ENG-008-06 Phase4;5"]

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    start_date = extract_start_milestone(xls)
    
    if start_date:
        st.success(f"📅 Начальная дата вехи найдена: **{start_date.strftime('%d.%m.%Y')}**")
        
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

            # Генерация файла
            excel_filename = create_true_gantt_chart(schedule_data)
            
            st.markdown("---")
            st.subheader("📊 Графический файл Excel сформирован")
            st.write("Файл содержит 2 вкладки: **«Реестр задач»** с исходными данными и **«Диаграмма Ганта»** с полноценным графическим объектом Excel.")

            # Кнопка скачивания Excel
            with open(excel_filename, "rb") as file:
                st.download_button(
                    label="📥 СКАЧАТЬ НАСТОЯЩУЮ ДИАГРАММУ ГАНТА В EXCEL (.XLSX)",
                    data=file,
                    file_name="SMS_Project_Gantt_Chart.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.error("Задачи не найдены в листах фаз.")
