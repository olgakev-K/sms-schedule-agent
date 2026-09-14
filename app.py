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

# Настройка страницы Streamlit
st.set_page_config(page_title="ИИ-Агент: SMS График", layout="wide")

st.title("🤖 ИИ-Агент: Генератор SMS-графика запуска в серийное производство")
st.markdown("""
**Задача агента:** Создать из файла-шаблона "PLANT_MASTER_SCHEDULE P25077" актуальный SMS-график проекта, 
учитывающий только рабочие дни и государственные праздники РФ.

**Выполняемые условия:**
1. Веха старта берется из вкладки `START PROJECT TOGF-ENG-007-02`.
2. Описание действий берется из столбца `DESCRIPTION` вкладок `Phase2`, `Phase 3`, `Phase4;5` в строгой последовательности.
3. Игнорируются все автоматические и плановые даты из исходного файла (`PLANNED START DATE`, `PLANNED END DATE` и т.д.).
4. Используется производственный календарь РФ (выходные и праздники).
""")

# Получение праздников РФ
@st.cache_data
def get_rf_holidays():
    holiday_dates = set()
    current_year = datetime.datetime.now().year
    # Используем библиотеку holidays для надежного получения всех праздников и переносов выходных
    ru_holidays = holidays.RU(years=[current_year - 1, current_year, current_year + 1, current_year + 2, current_year + 3])
    for d in ru_holidays.keys():
        holiday_dates.add(d)
    return holiday_dates

def is_business_day(date_val, holiday_dates):
    # Проверка на выходные (суббота=5, воскресенье=6) и государственные праздники
    if date_val.weekday() >= 5 or date_val in holiday_dates:
        return False
    return True

def get_next_business_day(date_val, holiday_dates):
    cur = date_val
    while not is_business_day(cur, holiday_dates):
        cur += datetime.timedelta(days=1)
    return cur

# 1. Извлечение даты вехи старта
def extract_start_milestone(xls):
    sheet_name = "START PROJECT TOGF-ENG-007-02"
    if sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                cell_val = df.iloc[r, c]
                if pd.notna(cell_val) and str(cell_val).strip().upper() not in ['N/A', 'NONE', 'CLOSED', 'DATE']:
                    dt = pd.to_datetime(cell_val, errors='coerce')
                    if pd.notna(dt) and 2020 < dt.year < 2035: # Фильтр адекватных годов
                        return dt.date()
    return None

# Генерация Excel-файла с компактной НЕДЕЛЬНОЙ Диаграммой Ганта
def generate_excel_with_gantt(schedule_data, start_date, total_days=60):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SMS Gantt Schedule"
    
    # Стили
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    gantt_header_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    task_bar_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
    
    # Шрифты: 10 для основных данных, 9 для столбцов диаграммы (по требованию)
    header_font = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    gantt_font = Font(name="Calibri", size=9, bold=False, color="000000") # Шрифт 9 для диаграммы
    gantt_header_font = Font(name="Calibri", size=9, bold=True, color="FFFFFF")
    regular_font = Font(name="Calibri", size=10)
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    # Базовые заголовки колонок
    headers = ["№", "Фаза проекта", "Описание работы (DESCRIPTION)", "Дата начала", "Дата окончания"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
    
    # Генерация списка недель для диаграммы
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
            'label': f"Н{week_start.isocalendar()[1]}\n{week_start.strftime('%d.%m')}"
        })
        current_date = week_end + datetime.timedelta(days=1)
        
    timeline_start_col = len(headers) + 1
    
    # Создание заголовков недель (применяем шрифт 9 и ширину 4)
    for i, week in enumerate(weeks_list):
        c_idx = timeline_start_col + i
        cell = ws.cell(row=1, column=c_idx, value=week['label'])
        cell.fill = gantt_header_fill
        cell.font = gantt_header_font
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border
        ws.column_dimensions[get_column_letter(c_idx)].width = 4  # Ширина столбцов диаграммы = 4

    # Заполнение данных и отрисовка Ганта
    for row_idx, item in enumerate(schedule_data, 2):
        ws.cell(row=row_idx, column=1, value=item['№']).font = regular_font
        ws.cell(row=row_idx, column=2, value=item['Фаза проекта']).font = regular_font
        ws.cell(row=row_idx, column=3, value=item['DESCRIPTION']).font = regular_font
        ws.cell(row=row_idx, column=4, value=item['Start'].strftime("%d.%m.%Y")).font = regular_font
        ws.cell(row=row_idx, column=5, value=item['Finish'].strftime("%d.%m.%Y")).font = regular_font
        
        # Отрисовка полос Ганта по неделям
        for i, week in enumerate(weeks_list):
            c_idx = timeline_start_col + i
            cell = ws.cell(row=row_idx, column=c_idx)
            cell.border = thin_border
            cell.font = gantt_font # Шрифт 9 во всех ячейках диаграммы
            
            # Если задача пересекается с неделей, закрашиваем ячейку
            if item['Start'] <= week['end'] and item['Finish'] >= week['start']:
                cell.fill = task_bar_fill

    # Фиксированная ширина текстовых колонок для аккуратного вида
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 25
    ws.column_dimensions['C'].width = 50
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 14
    
    # Закрепление верхней строки и первого столбца для удобства просмотра
    ws.freeze_panes = "F2"
    
    file_name = "SMS_Gantt_Schedule_AI_Agent.xlsx"
    wb.save(file_name)
    return file_name

# Загрузка файла
uploaded_file = st.file_uploader("Загрузите файл шаблона: PLANT_MASTER_SCHEDULE P25077.xlsx", type=["xlsx"])

# Строгая последовательность вкладок согласно условию 2
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
            st.success(f"✅ Веха проекта успешно извлечена: **{start_date.strftime('%d.%m.%Y')}** (из вкладки START PROJECT TOGF-ENG-007-02)")
            
            if st.button("🚀 Сформировать SMS-график с учетом праздников РФ", type="primary"):
                with st.spinner("ИИ-агент анализирует файл, игнорирует автоматические даты и строит календарный план..."):
                    tasks = []
                    
                    # Условие 2: Строгая последовательность вкладок
                    for sheet in target_phases:
                        if sheet in xls.sheet_names:
                            df = pd.read_excel(xls, sheet_name=sheet, header=None)
                            desc_col = None
                            
                            # Поиск столбца DESCRIPTION
                            for r in range(min(30, df.shape[0])):
                                for c in range(df.shape[1]):
                                    if str(df.iloc[r, c]).strip().upper() == 'DESCRIPTION':
                                        desc_col = c
                                        break
                                if desc_col is not None:
                                    break
                                    
                            if desc_col is not None:
                                # Извлекаем только описания, игнорируя заголовки и пустые строки
                                for val in df.iloc[10:, desc_col].dropna():
                                    val_str = str(val).strip()
                                    # Условие 3: Игнорируем строки, которые могут быть случайными датами или заголовками
                                    if val_str and val_str.upper() not in ['DESCRIPTION', 'N/A', 'NONE']:
                                        tasks.append({'Phase': sheet, 'DESCRIPTION': val_str})

                    if tasks:
                        df_tasks = pd.DataFrame(tasks)
                        holiday_dates = get_rf_holidays()
                        
                        current_date = start_date
                        schedule_data = []
                        gantt_plot_data = []
                        
                        # Построение календарного плана: каждая задача начинается в следующий рабочий день
                        for idx, row in df_tasks.iterrows():
                            current_date = get_next_business_day(current_date, holiday_dates)
                            
                            t_start = current_date
                            t_end = t_start # Длительность задачи по умолчанию 1 рабочий день (можно масштабировать при необходимости)
                            
                            schedule_data.append({
                                '№': idx + 1,
                                'Фаза проекта': row['Phase'],
                                'DESCRIPTION': row['DESCRIPTION'],
                                'Start': t_start,
                                'Finish': t_end
                            })
                            
                            task_name = f"{idx + 1}. {row['DESCRIPTION'][:45]}..." if len(row['DESCRIPTION']) > 45 else f"{idx + 1}. {row['DESCRIPTION']}"
                            gantt_plot_data.append({
                                'Task': task_name,
                                'Start': t_start,
                                'Finish': t_start + datetime.timedelta(days=1),
                                'Phase': row['Phase'],
                                'Full_Description': row['DESCRIPTION']
                            })
                            
                            # Переход к следующему рабочему дню для следующей задачи
                            current_date = get_next_business_day(current_date + datetime.timedelta(days=1), holiday_dates)
                        
                        # Веб-визуализация (недельный масштаб)
                        st.subheader("📊 Интерактивная Диаграмма Ганта (Недельный масштаб)")
                        fig = px.timeline(
                            pd.DataFrame(gantt_plot_data), 
                            x_start="Start", 
                            x_end="Finish", 
                            y="Task", 
                            color="Phase",
                            hover_data=["Full_Description"],
                            title="SMS-график проекта (с учетом выходных и праздников РФ)"
                        )
                        fig.update_yaxes(autorange="reversed")
                        
                        # Настройка осей для недельного отображения
                        fig.update_xaxes(
                            tickformat="%d.%m<br>(Н%W)",
                            tickangle=0,
                            dtick=7 * 24 * 60 * 60 * 1000,  # 7 дней в миллисекундах
                            ticklabelmode="period"
                        )
                        
                        chart_height = min(600, len(gantt_plot_data) * 30 + 150)
                        st.plotly_chart(fig, use_container_width=True, height=chart_height)
                        
                        # Генерация Excel
                        last_finish_date = schedule_data[-1]['Finish']
                        total_days_span = (last_finish_date - start_date).days + 21 # Запас в 3 недели
                        
                        excel_filename = generate_excel_with_gantt(schedule_data, start_date, total_days=total_days_span)
                        
                        st.subheader("📥 Выгрузка результатов")
                        with open(excel_filename, "rb") as f:
                            st.download_button(
                                label="📥 Скачать SMS-график в Excel (.xlsx)",
                                data=f,
                                file_name="SMS_Gantt_Schedule_AI_Agent.xlsx",
                                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                                type="primary"
                            )
                    else:
                        st.error("❌ Не удалось извлечь задачи из столбца DESCRIPTION в указанных вкладках.")
        else:
            st.error("❌ Не удалось найти дату старта на вкладке 'START PROJECT TOGF-ENG-007-02'. Проверьте файл.")
            
    except Exception as e:
        st.error(f"⚠️ Ошибка при обработке файла: {e}")
else:
    st.info("ℹ️ Ожидание загрузки файла шаблона PLANT_MASTER_SCHEDULE P25077.xlsx для начала работы ИИ-агента.")
