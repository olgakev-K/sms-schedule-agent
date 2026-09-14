import streamlit as st
import pandas as pd
import datetime
import requests
from bs4 import BeautifulSoup
import holidays
import plotly.express as px
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="SMS Schedule Agent - 3 Gantt Views", layout="wide")

st.title("🤖 ИИ-Агент: Генератор SMS-графика проекта (Сравнение 3 вариантов)")
st.write("Сравните три интерактивных варианта отображения календарного плана проекта и выберите наиболее удобный.")

# ---------------------------------------------------------
# 1. Получение праздников РФ
# ---------------------------------------------------------
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

# ---------------------------------------------------------
# 2. Извлечение вехи старта из Excel
# ---------------------------------------------------------
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

# ---------------------------------------------------------
# 3. Excel Генераторы для 3 Вариантов
# ---------------------------------------------------------

# Вариант 1: Встроенный графический чарт Excel (BarChart)
def generate_excel_variant1(schedule_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Gantt Chart View"
    
    headers = ["№", "Фаза проекта", "Описание работы", "Дата начала", "Длительность (дней)", "Смещение от старта"]
    ws.append(headers)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    base_date = schedule_data[0]['Start']
    
    for item in schedule_data:
        start_offset = (item['Start'] - base_date).days
        ws.append([
            item['№'],
            item['Фаза проекта'],
            item['DESCRIPTION'],
            item['Start'].strftime("%d.%m.%Y"),
            1,
            start_offset
        ])

    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 28
    ws.column_dimensions['C'].width = 45
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 18
    ws.column_dimensions['F'].width = 18

    chart = BarChart()
    chart.type = "barHoriz"
    chart.style = 10
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = "Вариант 1: Диаграмма Ганта (Штатный график Excel)"
    chart.height = max(10, len(schedule_data) * 0.5)
    chart.width = 18

    data = Reference(ws, min_col=5, min_row=1, max_col=6, max_row=len(schedule_data)+1)
    cats = Reference(ws, min_col=3, min_row=2, max_row=len(schedule_data)+1)
    
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    
    if len(chart.series) > 0:
        chart.series[0].graphicalProperties.solidFill = "FFFFFF"
        chart.series[0].graphicalProperties.line.solidFill = "FFFFFF"
        
    ws.add_chart(chart, "H2")
    
    filename = "SMS_Schedule_Variant1_Chart.xlsx"
    wb.save(filename)
    return filename

# Вариант 2: Иерархический график по фазам (Summary Milestones)
def generate_excel_variant2(phase_summary, schedule_data):
    wb = openpyxl.Workbook()
    
    # Лист 1: Сводные фазы
    ws_summary = wb.active
    ws_summary.title = "Свод по фазам (Phase Summary)"
    
    headers_sum = ["№", "Фаза проекта", "Кол-во задач", "Дата начала фазы", "Дата окончания фазы", "Длительность (дней)"]
    ws_summary.append(headers_sum)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    for col_num in range(1, len(headers_sum) + 1):
        cell = ws_summary.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")
        
    for idx, p in enumerate(phase_summary, 1):
        dur = (p['Finish'] - p['Start']).days + 1
        ws_summary.append([
            idx,
            p['Phase'],
            p['TaskCount'],
            p['Start'].strftime("%d.%m.%Y"),
            p['Finish'].strftime("%d.%m.%Y"),
            dur
        ])
        
    for col in ['A', 'B', 'C', 'D', 'E', 'F']:
        ws_summary.column_dimensions[col].width = 22
        
    # График фаз
    chart = BarChart()
    chart.type = "barHoriz"
    chart.title = "Вариант 2: Сводный график по Фазам Проекта"
    chart.height = 8
    chart.width = 16
    
    data = Reference(ws_summary, min_col=6, min_row=1, max_row=len(phase_summary)+1)
    cats = Reference(ws_summary, min_col=2, min_row=2, max_row=len(phase_summary)+1)
    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)
    ws_summary.add_chart(chart, "H2")
    
    # Лист 2: Детализация задач
    ws_detail = wb.create_sheet(title="Полный список задач")
    ws_detail.append(["№", "Фаза проекта", "Описание работы (DESCRIPTION)", "Дата"])
    for col_num in range(1, 5):
        cell = ws_detail.cell(row=1, column=col_num)
        cell.fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
        cell.font = header_font
        
    for item in schedule_data:
        ws_detail.append([item['№'], item['Фаза проекта'], item['DESCRIPTION'], item['Start'].strftime("%d.%m.%Y")])
        
    ws_detail.column_dimensions['A'].width = 6
    ws_detail.column_dimensions['B'].width = 28
    ws_detail.column_dimensions['C'].width = 50
    ws_detail.column_dimensions['D'].width = 14
    
    filename = "SMS_Schedule_Variant2_Phases.xlsx"
    wb.save(filename)
    return filename

# Вариант 3: Недельная матрица (Weekly Matrix)
def generate_excel_variant3(schedule_data, start_date):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Weekly Matrix Schedule"
    
    headers = ["№", "Фаза проекта", "Описание работы", "Дата"]
    
    # Расчет недель
    min_date = schedule_data[0]['Start']
    max_date = schedule_data[-1]['Start']
    
    # Получаем список уникальных недель
    weeks = []
    curr = min_date
    while curr <= max_date + datetime.timedelta(days=7):
        year, week_num, _ = curr.isocalendar()
        w_label = f"W{week_num} ({curr.strftime('%d.%m')})"
        if w_label not in [w['label'] for w in weeks]:
            weeks.append({'year': year, 'week': week_num, 'label': w_label})
        curr += datetime.timedelta(days=7)

    all_headers = headers + [w['label'] for w in weeks]
    ws.append(all_headers)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    week_header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    task_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    
    for c_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=c_idx)
        cell.fill = header_fill
        cell.font = header_font
        
    for c_idx in range(len(headers) + 1, len(all_headers) + 1):
        cell = ws.cell(row=1, column=c_idx)
        cell.fill = week_header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center")
        ws.column_dimensions[get_column_letter(c_idx)].width = 12

    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 40
    ws.column_dimensions['D'].width = 12

    # Заполнение таблицы
    for r_idx, item in enumerate(schedule_data, 2):
        ws.cell(row=r_idx, column=1, value=item['№'])
        ws.cell(row=r_idx, column=2, value=item['Фаза проекта'])
        ws.cell(row=r_idx, column=3, value=item['DESCRIPTION'])
        ws.cell(row=r_idx, column=4, value=item['Start'].strftime("%d.%m.%Y"))
        
        task_year, task_week, _ = item['Start'].isocalendar()
        
        for w_idx, w in enumerate(weeks, len(headers) + 1):
            cell = ws.cell(row=r_idx, column=w_idx)
            if w['year'] == task_year and w['week'] == task_week:
                cell.fill = task_fill
                cell.value = "✓"
                cell.alignment = Alignment(horizontal="center")
                cell.font = Font(bold=True, color="FFFFFF")

    filename = "SMS_Schedule_Variant3_Weekly.xlsx"
    wb.save(filename)
    return filename

# ---------------------------------------------------------
# 4. Streamlit Интерфейс
# ---------------------------------------------------------

uploaded_file = st.file_uploader("Загрузите Excel-файл проекта (PLANT_MASTER_SCHEDULE P25077.xlsx)", type=["xlsx"])
target_phases = ["PMSPR TOGF-ENG-008-06 Phase2", "PMSPR TOGF-ENG-008-06 Phase 3", "PMSPR TOGF-ENG-008-06 Phase4;5"]

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    start_date = extract_start_milestone(xls)
    
    if start_date:
        st.success(f"📅 Начальная дата вехи: **{start_date.strftime('%d.%m.%Y')}**")
        
        if st.button("🚀 Рассчитать и сравнить 3 варианта графика"):
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
                        'Start': current_date
                    })
                    current_date = get_next_business_day(current_date + datetime.timedelta(days=1), holiday_dates)

                # Генерация данных по фазам (для Варианта 2)
                phase_summary_dict = {}
                for item in schedule_data:
                    ph = item['Фаза проекта']
                    if ph not in phase_summary_dict:
                        phase_summary_dict[ph] = {'Phase': ph, 'Start': item['Start'], 'Finish': item['Start'], 'TaskCount': 0}
                    phase_summary_dict[ph]['Finish'] = item['Start']
                    phase_summary_dict[ph]['TaskCount'] += 1
                phase_summary = list(phase_summary_dict.values())

                # Генерация Excel файлов
                ex1 = generate_excel_variant1(schedule_data)
                ex2 = generate_excel_variant2(phase_summary, schedule_data)
                ex3 = generate_excel_variant3(schedule_data, start_date)

                # ---------------------------------------------------------
                # Вкладки для сравнения 3 вариантов
                # ---------------------------------------------------------
                tab1, tab2, tab3 = st.tabs([
                    "📈 Вариант 1: Штатный Chart Excel", 
                    "📊 Вариант 2: Свод по Фазам", 
                    "📅 Вариант 3: Недельная Матрица"
                ])

                # ТАБ 1
                with tab1:
                    st.subheader("Вариант 1: Графический блок в Excel (BarChart)")
                    st.info("💡 **Особенности:** В Excel создается чистый графический объект без растягивания сетки ячеек. Таблица слева остаётся компактной и читабельной.")
                    
                    gantt_plot = []
                    for item in schedule_data:
                        t_name = f"{item['№']}. {item['DESCRIPTION'][:50]}..." if len(item['DESCRIPTION']) > 50 else f"{item['№']}. {item['DESCRIPTION']}"
                        gantt_plot.append({
                            'Task': t_name,
                            'Start': item['Start'],
                            'Finish': item['Start'] + datetime.timedelta(days=1),
                            'Phase': item['Фаза проекта']
                        })
                    fig1 = px.timeline(pd.DataFrame(gantt_plot), x_start="Start", x_end="Finish", y="Task", color="Phase", title="Интерактивный вид в приложении")
                    fig1.update_yaxes(autorange="reversed")
                    st.plotly_chart(fig1, use_container_width=True)
                    
                    with open(ex1, "rb") as f:
                        st.download_button("📥 Скачать Excel (Вариант 1)", f, file_name=ex1)

                # ТАБ 2
                with tab2:
                    st.subheader("Вариант 2: Сводный график по фазам проекта")
                    st.info("💡 **Особенности:** Отображаются только крупные этапы (Phase 2, Phase 3 и т.д.). Идеально для руководителей и быстрых отчетов.")
                    
                    df_phase = pd.DataFrame(phase_summary)
                    df_phase['Start_str'] = df_phase['Start'].apply(lambda x: x.strftime('%d.%m.%Y'))
                    df_phase['Finish_str'] = df_phase['Finish'].apply(lambda x: x.strftime('%d.%m.%Y'))
                    
                    st.table(df_phase[['Phase', 'TaskCount', 'Start_str', 'Finish_str']])
                    
                    fig2 = px.timeline(df_phase, x_start="Start", x_end="Finish", y="Phase", color="Phase", title="Сроки фаз проекта")
                    st.plotly_chart(fig2, use_container_width=True)
                    
                    with open(ex2, "rb") as f:
                        st.download_button("📥 Скачать Excel (Вариант 2)", f, file_name=ex2)

                # ТАБ 3
                with tab3:
                    st.subheader("Вариант 3: Недельная матрица (Weekly View)")
                    st.info("💡 **Особенности:** Шкала разбита по неделям (W41, W42, W43...). Вся диаграмма легко помещается на одном экране без бесконечной прокрутки.")
                    
                    df_week = pd.DataFrame(schedule_data)
                    df_week['Дата'] = df_week['Start'].apply(lambda x: x.strftime('%d.%m.%Y'))
                    df_week['Неделя'] = df_week['Start'].apply(lambda x: f"W{x.isocalendar()[1]}")
                    
                    st.dataframe(df_week[['№', 'Фаза проекта', 'DESCRIPTION', 'Дата', 'Неделя']], use_container_width=True)
                    
                    with open(ex3, "rb") as f:
                        st.download_button("📥 Скачать Excel (Вариант 3)", f, file_name=ex3)
            else:
                st.error("Не удалось извлечь задачи из столбцов DESCRIPTION.")
    else:
        st.error("Не удалось извлечь начальную дату.")
