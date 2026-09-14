import streamlit as st
import pandas as pd
import datetime
import holidays
import plotly.express as px
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import PatternFill, Font, Alignment

st.set_page_config(page_title="SMS Schedule Agent — Weekly Gantt Chart", layout="wide")

st.title("📊 ИИ-Агент: Недельная Диаграмма Ганта проекта")
st.write("Формирование недельного графического плана работ и выгрузка отчета в Excel")

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

def create_excel_with_weekly_gantt(schedule_data, project_start_date):
    wb = openpyxl.Workbook()
    
    # 1. Лист «Реестр задач (по неделям)»
    ws_data = wb.active
    ws_data.title = "Реестр задач (Недели)"
    ws_data.views.sheetView[0].showGridLines = True

    headers = [
        "№", 
        "Фаза проекта", 
        "Описание работы (DESCRIPTION)", 
        "Дата начала", 
        "Дата финиша", 
        "Номер недели (ISO)", 
        "Смещение (недель)", 
        "Длительность (недель)"
    ]
    ws_data.append(headers)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    
    for col_idx in range(1, len(headers) + 1):
        cell = ws_data.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    # Приводим старт проекта к началу недели (понедельник)
    project_start_monday = project_start_date - datetime.timedelta(days=project_start_date.weekday())

    for item in schedule_data:
        # Расчет сдвига и длительности в неделях
        start_monday = item['Start'] - datetime.timedelta(days=item['Start'].weekday())
        finish_sunday = item['Finish'] + datetime.timedelta(days=(6 - item['Finish'].weekday()))
        
        offset_weeks = max(0, (start_monday - project_start_monday).days // 7)
        duration_weeks = max(1, ((finish_sunday - start_monday).days + 1) // 7)
        week_label = f"W{item['Start'].isocalendar()[1]} ({item['Start'].strftime('%d.%m')})"

        ws_data.append([
            item['№'],
            item['Фаза проекта'],
            f"{item['№']}. {item['DESCRIPTION']}",
            item['Start'].strftime("%d.%m.%Y"),
            item['Finish'].strftime("%d.%m.%Y"),
            week_label,
            offset_weeks,
            duration_weeks
        ])

    ws_data.column_dimensions['A'].width = 6
    ws_data.column_dimensions['B'].width = 28
    ws_data.column_dimensions['C'].width = 50
    ws_data.column_dimensions['D'].width = 14
    ws_data.column_dimensions['E'].width = 14
    ws_data.column_dimensions['F'].width = 18
    ws_data.column_dimensions['G'].width = 18
    ws_data.column_dimensions['H'].width = 20

    # 2. Лист «График Ганта (Недельный)»
    ws_chart = wb.create_sheet(title="График Ганта (Недели)")
    ws_chart.views.sheetView[0].showGridLines = True

    chart = BarChart()
    chart.type = "bar"
    chart.dir = "bar"
    chart.style = 13
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = f"Недельная Диаграмма Ганта проекта (Старт: {project_start_date.strftime('%d.%m.%Y')})"
    chart.x_axis.title = "Шкала времени (недели)"
    chart.y_axis.title = "Задачи"
    chart.height = max(14, len(schedule_data) * 0.8)
    chart.width = 24

    # Столбцы 7 и 8: Смещение (недель) и Длительность (недель)
    data = Reference(ws_data, min_col=7, min_row=1, max_col=8, max_row=len(schedule_data) + 1)
    cats = Reference(ws_data, min_col=3, min_row=2, max_row=len(schedule_data) + 1)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    # Делаем прозрачным ряд со сдвигом недели
    if len(chart.series) > 0:
        chart.series[0].graphicalProperties.solidFill = "FFFFFF"
        chart.series[0].graphicalProperties.line.solidFill = "FFFFFF"

    ws_chart.add_chart(chart, "B2")

    filename = "SMS_Project_Weekly_Gantt.xlsx"
    wb.save(filename)
    return filename

# --- ИНТЕРФЕЙС STREAMLIT ---
uploaded_file = st.file_uploader("Загрузите Excel-файл проекта (PLANT_MASTER_SCHEDULE P25077.xlsx)", type=["xlsx"])
target_phases = ["PMSPR TOGF-ENG-008-06 Phase2", "PMSPR TOGF-ENG-008-06 Phase 3", "PMSPR TOGF-ENG-008-06 Phase4;5"]

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    start_date = extract_start_milestone(xls)
    
    if start_date:
        st.success(f"📅 Дата старта вехи: **{start_date.strftime('%d.%m.%Y')}**")
        
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
                finish_date = current_date + datetime.timedelta(days=1)
                
                # Привязка к понедельникам недель
                week_num = current_date.isocalendar()[1]
                week_start_monday = current_date - datetime.timedelta(days=current_date.weekday())
                
                schedule_data.append({
                    '№': idx + 1,
                    'Фаза проекта': row['Phase'],
                    'DESCRIPTION': f"{idx + 1}. {row['DESCRIPTION']}",
                    'Start': current_date,
                    'Finish': finish_date,
                    'Week_Num': week_num,
                    'Week_Label': f"Неделя {week_num} ({week_start_monday.strftime('%d.%m')})"
                })
                current_date = get_next_business_day(finish_date, holiday_dates)

            df_sched = pd.DataFrame(schedule_data)

            # --- ВЕБ-ДИАГРАММА ГАНТА ПО НЕДЕЛЯМ (Plotly) ---
            st.markdown("### 📈 График Ганта по неделям")
            
            fig = px.timeline(
                df_sched,
                x_start="Start",
                x_end="Finish",
                y="DESCRIPTION",
                color="Фаза проекта",
                hover_data=["№", "Start", "Finish", "Week_Label"],
                title="Недельный план-график выполнения работ"
            )
            
            fig.update_yaxes(autorange="reversed", title="Задачи / Описание работ")
            
            # Переключение оси времени на недельный шаг (M1 = 1 неделя)
            fig.update_xaxes(
                title="Шкала времени (недели)",
                dtick="M1",  
                tickformat="%d.%m\n(W%V)", # Число.Месяц + Номер недели ISO
                rangeslider=dict(visible=True)
            )
            
            fig.update_layout(
                height=max(600, len(schedule_data) * 25),
                legend_title_text="Фазы проекта",
                font=dict(size=12)
            )

            st.plotly_chart(fig, use_container_width=True)

            # --- ВЫГРУЗКА В EXCEL ---
            excel_file = create_excel_with_weekly_gantt(schedule_data, start_date)
            
            st.markdown("---")
            with open(excel_file, "rb") as f:
                st.download_button(
                    label="📥 СКАЧАТЬ НЕДЕЛЬНУЮ ДИАГРАММУ ГАНТА В EXCEL (.XLSX)",
                    data=f,
                    file_name="SMS_Weekly_Gantt_Schedule.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.error("Задачи не найдены в исходных листах.")
