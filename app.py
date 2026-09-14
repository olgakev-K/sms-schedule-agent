import streamlit as st
import pandas as pd
import datetime
import holidays
import plotly.express as px
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

st.set_page_config(
    page_title="SMS Schedule Agent — Weekly Gantt", 
    layout="wide"
)

st.title("📊 ИИ-Агент: SMS График (Недельная Диаграмма Ганта)")

@st.cache_data
def get_rf_holidays():
    current_year = datetime.datetime.now().year
    ru_holidays = holidays.RU(
        years=[current_year - 2, current_year - 1, current_year, current_year + 1, current_year + 2]
    )
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

def create_excel_with_blue_gantt(schedule_data, project_start_date):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SMS График Ганта"
    ws.views.sheetView[0].showGridLines = True

    headers = [
        "№", 
        "Фаза проекта", 
        "Описание работы", 
        "Дата начала", 
        "Дата финиша", 
        "Неделя", 
        "Смещение (недель)", 
        "Длительность (недель)"
    ]
    ws.append(headers)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )
    
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    project_start_monday = project_start_date - datetime.timedelta(days=project_start_date.weekday())

    for row_idx, item in enumerate(schedule_data, start=2):
        start_monday = item['Start'] - datetime.timedelta(days=item['Start'].weekday())
        finish_sunday = item['Finish'] + datetime.timedelta(days=(6 - item['Finish'].weekday()))
        
        offset_weeks = max(0, (start_monday - project_start_monday).days // 7)
        duration_weeks = max(1, ((finish_sunday - start_monday).days + 1) // 7)
        week_label = f"W{item['Start'].isocalendar()[1]}"

        ws.append([
            item['№'],
            item['Фаза проекта'],
            item['DESCRIPTION'],
            item['Start'].strftime("%d.%m.%Y"),
            item['Finish'].strftime("%d.%m.%Y"),
            week_label,
            offset_weeks,
            duration_weeks
        ])

        row_fill = PatternFill(start_color="F2F7FA" if row_idx % 2 == 0 else "FFFFFF", fill_type="solid")
        for col_i in range(1, len(headers) + 1):
            c = ws.cell(row=row_idx, column=col_i)
            c.fill = row_fill
            c.border = thin_border

    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 24
    ws.column_dimensions['C'].width = 42
    ws.column_dimensions['D'].width = 13
    ws.column_dimensions['E'].width = 13
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 18
    ws.column_dimensions['H'].width = 20

    chart = BarChart()
    chart.type = "bar"
    chart.dir = "bar"
    chart.style = 13
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = f"Недельная Диаграмма Ганта (Старт: {project_start_date.strftime('%d.%m.%Y')})"
    chart.x_axis.title = "Недели"
    chart.height = max(12, len(schedule_data) * 0.75)
    chart.width = 18

    data = Reference(ws, min_col=7, min_row=1, max_col=8, max_row=len(schedule_data) + 1)
    cats = Reference(ws, min_col=3, min_row=2, max_row=len(schedule_data) + 1)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    if len(chart.series) > 1:
        chart.series[0].graphicalProperties.solidFill = "FFFFFF"
        chart.series[0].graphicalProperties.line.solidFill = "FFFFFF"
        chart.series[1].graphicalProperties.solidFill = "1F4E78"
        chart.series[1].graphicalProperties.line.solidFill = "1F4E78"

    ws.add_chart(chart, "J1")

    filename = "SMS_Weekly_Gantt_Blue_Schedule.xlsx"
    wb.save(filename)
    return filename

# --- ИНТЕРФЕЙС ---
uploaded_file = st.file_uploader("Загрузите файл проекта (.xlsx)", type=["xlsx"])
target_phases = ["PMSPR TOGF-ENG-008-06 Phase2", "PMSPR TOGF-ENG-008-06 Phase 3", "PMSPR TOGF-ENG-008-06 Phase4;5"]

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    start_date = extract_start_milestone(xls)
    
    if start_date:
        st.success(f"📅 Дата старта вехи проекта: **{start_date.strftime('%d.%m.%Y')}**")
        
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
                
                # Привязка старта и конца строго к рабочей неделе (без дневных дыр)
                week_start_monday = current_date - datetime.timedelta(days=current_date.weekday())
                week_end_sunday = week_start_monday + datetime.timedelta(days=6)
                week_num = current_date.isocalendar()[1]
                
                schedule_data.append({
                    '№': idx + 1,
                    'Фаза проекта': row['Phase'],
                    'DESCRIPTION': f"{idx + 1}. {row['DESCRIPTION']}",
                    'Start': week_start_monday,
                    'Finish': week_end_sunday,
                    'Week_Num': week_num,
                    'Week_Label': f"W{week_num} ({week_start_monday.strftime('%d.%m')})"
                })
                current_date = get_next_business_day(current_date + datetime.timedelta(days=1), holiday_dates)

            df_sched = pd.DataFrame(schedule_data)

            st.markdown("---")
            st.markdown("### 📈 Автоматический SMS График (Недельная Диаграмма Ганта)")

            # Синяя гамма
            blue_shades = ["#1F4E78", "#2F5597", "#41719C", "#5B9BD5", "#8EA9DB"]

            fig = px.timeline(
                df_sched,
                x_start="Start",
                x_end="Finish",
                y="DESCRIPTION",
                color="Фаза проекта",
                hover_data=["№", "Week_Label"],
                color_discrete_sequence=blue_shades,
                title="План-график выполнения работ (по неделям)"
            )
            
            fig.update_yaxes(autorange="reversed", title="")
            
            # Убираем отображение пустых дней, задаем недельный шаг
            fig.update_xaxes(
                title="Недели проекта",
                dtick="M1", # Недельный шаг
                tickformat="%d.%m\n(W%V)",
                showgrid=True,
                gridcolor="#E2E8F0"
            )
            
            fig.update_layout(
                height=max(500, len(schedule_data) * 26),
                plot_bgcolor="#FFFFFF",
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.01,
                    xanchor="right",
                    x=1,
                    title_text="Фазы:"
                ),
                font=dict(size=12)
            )

            # ОТОБРАЖЕНИЕ ТОЛЬКО ДИАГРАММЫ ГАНТА В ОТВЕТЕ
            st.plotly_chart(fig, use_container_width=True)

            # КНОПКА СКАЧИВАНИЯ
            excel_file = create_excel_with_blue_gantt(schedule_data, start_date)
            
            st.markdown("---")
            with open(excel_file, "rb") as f:
                st.download_button(
                    label="📥 СКАЧАТЬ СФОРМИРОВАННЫЙ SMS ГРАФИК (EXCEL .XLSX)",
                    data=f,
                    file_name="SMS_Weekly_Gantt_Blue_Schedule.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.error("Задачи не найдены.")
