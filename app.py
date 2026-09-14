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

st.set_page_config(page_title="SMS Schedule Agent", layout="wide")

st.title("🤖 ИИ-Агент: Генератор SMS-графика проекта")
st.write("Расчет календарного плана и выгрузка Excel с визуально компактной недельной Диаграммой Ганта")

# 4. Праздники РФ
@st.cache_data
def get_rf_holidays():
    url = "https://www.consultant.ru/law/ref/calendar/proizvodstvennye/"
    headers = {"User-Agent": "Mozilla/5.0"}
    holiday_dates = set()
    
    try:
        response = requests.get(url, headers=headers, timeout=5)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            for td in soup.find_all('td', class_=['holiday', 'work_short']):
                pass # Логика парсинга оставлена для совместимости, основная нагрузка на holidays.RU
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

# 1. Извлечение вехи/даты из "START PROJECT TOGF-ENG-007-02"
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

# Функция генерации компактного Excel-файла с НЕДЕЛЬНОЙ Диаграммой Ганта
def generate_excel_with_gantt(schedule_data, start_date, total_days=30):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SMS Gantt Schedule"
    
    # Стили
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    gantt_header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    task_bar_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    regular_font = Font(name="Calibri", size=10)
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    # Базовые заголовки колонок
    headers = ["№", "Фаза", "Описание работы", "Начало", "Окончание"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
    
    # Генерация списка недель вместо дней
    weeks_list = []
    current_date = start_date
    end_date = start_date + datetime.timedelta(days=total_days)
    
    while current_date <= end_date:
        week_start = current_date
        week_end = week_start + datetime.timedelta(days=6)
        if week_end > end_date:
            week_end = end_date
            
        weeks_list.append({
            'start': week_start,
            'end': week_end,
            'label': f"Нед.\n{week_start.strftime('%d.%m')}\n-\n{week_end.strftime('%d.%m')}"
        })
        current_date = week_end + datetime.timedelta(days=1)
        
    timeline_start_col = len(headers) + 1
    
    # Создание заголовков недель
    for i, week in enumerate(weeks_list):
        c_idx = timeline_start_col + i
        cell = ws.cell(row=1, column=c_idx, value=week['label'])
        cell.fill = gantt_header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(c_idx)].width = 10  # Компактная ширина колонки недели

    # Заполнение данных и отрисовка Ганта по неделям
    for row_idx, item in enumerate(schedule_data, 2):
        ws.cell(row=row_idx, column=1, value=item['№']).font = regular_font
        ws.cell(row=row_idx, column=2, value=item['Фаза проекта']).font = regular_font
        ws.cell(row=row_idx, column=3, value=item['DESCRIPTION']).font = regular_font
        ws.cell(row=row_idx, column=4, value=item['Start'].strftime("%d.%m.%Y")).font = regular_font
        ws.cell(row=row_idx, column=5, value=item['Finish'].strftime("%d.%m.%Y")).font = regular_font
        
        # Отрисовка полос Ганта: если задача пересекается с неделей, закрашиваем ячейку
        for i, week in enumerate(weeks_list):
            c_idx = timeline_start_col + i
            cell = ws.cell(row=row_idx, column=c_idx)
            cell.border = thin_border
            
            # Проверка пересечения интервалов: [Start, Finish] и [week_start, week_end]
            if item['Start'] <= week['end'] and item['Finish'] >= week['start']:
                cell.fill = task_bar_fill

    # Авто-ширина текстовых колонок
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 40
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 12
    
    file_name = "SMS_Gantt_Schedule_Weekly.xlsx"
    wb.save(file_name)
    return file_name

uploaded_file = st.file_uploader("Загрузите файл (PLANT_MASTER_SCHEDULE P25077.xlsx)", type=["xlsx"])

target_phases = [
    "PMSPR TOGF-ENG-008-06 Phase2",
    "PMSPR TOGF-ENG-008-06 Phase 3",
    "PMSPR TOGF-ENG-008-06 Phase4;5"
]

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        start_date = extract_start_milestone(xls)
        
        if start_date:
            st.success(f"📅 Дата вехи извлечена из листа 'START PROJECT TOGF-ENG-007-02': **{start_date.strftime('%d.%m.%Y')}**")
            
            if st.button("🚀 Сформировать SMS-график и Excel с недельным Гантом"):
                with st.spinner("Расчет календарного плана и построение компактной диаграммы..."):
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
                        df_tasks = pd.DataFrame(tasks)
                        holiday_dates = get_rf_holidays()
                        
                        current_date = start_date
                        schedule_data = []
                        gantt_plot_data = []
                        
                        for idx, row in df_tasks.iterrows():
                            current_date = get_next_business_day(current_date, holiday_dates)
                            
                            t_start = current_date
                            t_end = t_start
                            
                            schedule_data.append({
                                '№': idx + 1,
                                'Фаза проекта': row['Phase'],
                                'DESCRIPTION': row['DESCRIPTION'],
                                'Start': t_start,
                                'Finish': t_end
                            })
                            
                            task_name = f"{idx + 1}. {row['DESCRIPTION'][:50]}..." if len(row['DESCRIPTION']) > 50 else f"{idx + 1}. {row['DESCRIPTION']}"
                            gantt_plot_data.append({
                                'Task': task_name,
                                'Start': t_start,
                                'Finish': t_start + datetime.timedelta(days=1), # Минимальная длительность для отображения
                                'Phase': row['Phase'],
                                'Full_Description': row['DESCRIPTION']
                            })
                            
                            current_date = get_next_business_day(current_date + datetime.timedelta(days=1), holiday_dates)
                        
                        # Отображение Plotly Ганта на веб-странице (недельный масштаб)
                        st.subheader("📊 Интерактивная Диаграмма Ганта (Недельный масштаб)")
                        fig = px.timeline(
                            pd.DataFrame(gantt_plot_data), 
                            x_start="Start", 
                            x_end="Finish", 
                            y="Task", 
                            color="Phase",
                            hover_data=["Full_Description"],
                            title="SMS-график проекта"
                        )
                        fig.update_yaxes(autorange="reversed")
                        
                        # Настройка осей для недельного отображения и компактности
                        fig.update_xaxes(
                            tickformat="%d.%m<br>(Нед. %W)",
                            tickangle=0,
                            dtick=7 * 24 * 60 * 60 * 1000,  # Интервал тиков: 7 дней в миллисекундах
                            ticklabelmode="period"
                        )
                        
                        # Динамическая, но компактная высота графика
                        chart_height = min(600, len(gantt_plot_data) * 35 + 150)
                        st.plotly_chart(fig, use_container_width=True, height=chart_height)
                        
                        # Генерация файла Excel с графической диаграммой Ганта
                        last_finish_date = schedule_data[-1]['Finish']
                        total_days_span = (last_finish_date - start_date).days + 15
                        
                        excel_filename = generate_excel_with_gantt(schedule_data, start_date, total_days=total_days_span)
                        
                        st.subheader("📥 Выгрузка результатов")
                        with open(excel_filename, "rb") as f:
                            st.download_button(
                                label="📥 Скачать Excel с недельной диаграммой Ганта (.xlsx)",
                                data=f,
                                file_name="SMS_Gantt_Schedule_Weekly.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                            )
                    else:
                        st.error("❌ Не удалось извлечь задачи из столбцов DESCRIPTION.")
        else:
            st.error("❌ Не удалось найти дату старта на вкладке 'START PROJECT TOGF-ENG-007-02'.")
            
    except Exception as e:
        st.error(f"Ошибка при обработке файла: {e}")
else:
    st.info("ℹ️ Для запуска расчета загрузите Excel-файл.")
