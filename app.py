import streamlit as st
import pandas as pd
import datetime
import requests
import holidays
import plotly.express as px
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="SMS Schedule Agent", layout="wide")

st.title("🤖 ИИ-Агент: Генератор интерактивного графика проекта")
st.write("Наглядная визуализация большого объема задач с возможностью масштабирования и выгрузки")

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

def generate_tree_excel(df_schedule):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "Сводный график MS Project style"
    ws.views.sheetView[0].showGridLines = True

    # Header
    headers = ["№", "Фаза проекта", "Описание работы (DESCRIPTION)", "Дата начала", "Дата завершения", "Рабочих дней", "Неделя"]
    ws.append(headers)
    
    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    for col_num in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_num)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    for r_idx, row in df_schedule.iterrows():
        curr_row = r_idx + 2
        ws.cell(row=curr_row, column=1, value=row['№']).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_row, column=2, value=row['Фаза проекта'])
        ws.cell(row=curr_row, column=3, value=row['DESCRIPTION'])
        ws.cell(row=curr_row, column=4, value=row['Start'].strftime("%d.%m.%Y")).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_row, column=5, value=row['Finish'].strftime("%d.%m.%Y")).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_row, column=6, value=1).alignment = Alignment(horizontal="center")
        ws.cell(row=curr_row, column=7, value=f"W{row['Start'].isocalendar()[1]}").alignment = Alignment(horizontal="center")

        for c_idx in range(1, len(headers) + 1):
            ws.cell(row=curr_row, column=c_idx).border = thin_border

    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 30
    ws.column_dimensions['C'].width = 55
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 14
    ws.column_dimensions['F'].width = 14
    ws.column_dimensions['G'].width = 10

    filename = "MS_Project_Style_Schedule.xlsx"
    wb.save(filename)
    return filename

uploaded_file = st.file_uploader("Загрузите Excel-файл проекта (PLANT_MASTER_SCHEDULE P25077.xlsx)", type=["xlsx"])
target_phases = ["PMSPR TOGF-ENG-008-06 Phase2", "PMSPR TOGF-ENG-008-06 Phase 3", "PMSPR TOGF-ENG-008-06 Phase4;5"]

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    start_date = extract_start_milestone(xls)
    
    if start_date:
        st.success(f"📅 Начальная дата вехи: **{start_date.strftime('%d.%m.%Y')}**")
        
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
            schedule_list = []
            
            for idx, row in enumerate(tasks):
                current_date = get_next_business_day(current_date, holiday_dates)
                # Для корректного отображения полос на Plotly график добавим продолжительность в 1 рабочий день
                finish_date = current_date + datetime.timedelta(days=1)
                schedule_list.append({
                    '№': idx + 1,
                    'Фаза проекта': row['Phase'],
                    'DESCRIPTION': f"{idx + 1}. {row['DESCRIPTION']}",
                    'Start': current_date,
                    'Finish': finish_date
                })
                current_date = get_next_business_day(finish_date, holiday_dates)

            df_schedule = pd.DataFrame(schedule_list)

            # Переключатель представлений
            view_option = st.radio(
                "Выберите формат отображения графика:",
                ["📊 1. Интерактивная диаграмма Ганта (Plotly со слайдером времени)", 
                 "📋 2. Сводная структура (MS Project style + Excel)", 
                 "🗓️ 3. Месячный обзор (Heatmap Overview)"]
            )

            if "1. Интерактивная" in view_option:
                st.write("### 🔍 Интерактивный график (подходит для сотен задач)")
                st.info("💡 **Как пользователем управлять:** Нажимайте и удерживайте левую кнопку мыши для приближения конкретного отрезка. Внизу расположена полоса масштабирования (Range Slider).")
                
                # Создаем Plotly Timeline
                fig = px.timeline(
                    df_schedule, 
                    x_start="Start", 
                    x_end="Finish", 
                    y="DESCRIPTION", 
                    color="Фаза проекта",
                    hover_data=["№", "Start", "Finish"],
                    title="График выполнения работ по проекту"
                )
                
                # Порядок задач сверху вниз
                fig.update_yaxes(autorange="reversed")
                
                # Настройка интерактивных элементов
                fig.update_layout(
                    height=max(500, len(df_schedule) * 22),
                    xaxis=dict(
                        rangeslider=dict(visible=True),
                        type="date"
                    ),
                    legend_title_text="Фазы проекта",
                    font=dict(size=11)
                )
                
                st.plotly_chart(fig, use_container_width=True)

            elif "2. Сводная структура" in view_option:
                st.write("### 📋 Табличный вид с разбивкой по неделям")
                st.dataframe(df_schedule[['№', 'Фаза проекта', 'DESCRIPTION', 'Start', 'Finish']], use_container_width=True)
                
                excel_file = generate_tree_excel(df_schedule)
                with open(excel_file, "rb") as f:
                    st.download_button("📥 Скачать аккуратный Excel-справочник", f, file_name="Project_Schedule.xlsx")

            elif "3. Месячный обзор" in view_option:
                st.write("### 🗓️ Распределение задач по месяцам")
                df_schedule['Месяц'] = df_schedule['Start'].apply(lambda x: x.strftime('%Y-%m (%B)'))
                monthly_summary = df_schedule.groupby(['Фаза проекта', 'Месяц']).size().unstack(fill_value=0)
                st.bar_chart(monthly_summary)
        else:
            st.error("Задачи не найдены в файле.")
