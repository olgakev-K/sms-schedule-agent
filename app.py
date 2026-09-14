import streamlit as st
import pandas as pd
import datetime
import holidays
import plotly.express as px
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

st.set_page_config(page_title="SMS Schedule Agent — OnePage Blue Gantt", layout="wide")

st.title("📊 ИИ-Агент: Недельный план-график проекта (Синяя схема)")

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

def create_excel_one_page(schedule_data, project_start_date):
    wb = openpyxl.Workbook()
    
    # Реестр и Диаграмма Ганта на ОДНОМ листе
    ws = wb.active
    ws.title = "План и Гант"
    ws.views.sheetView[0].showGridLines = True

    headers = [
        "№", 
        "Фаза", 
        "Описание работы", 
        "Старт", 
        "Финиш", 
        "Неделя", 
        "Смещение (недель)", 
        "Длительность (недель)"
    ]
    ws.append(headers)

    # Синее оформление заголовка
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

        # Чередование строк с легким синим оттенком
        row_fill = PatternFill(start_color="F2F7FA" if row_idx % 2 == 0 else "FFFFFF", fill_type="solid")
        for col_i in range(1, len(headers) + 1):
            c = ws.cell(row=row_idx, column=col_i)
            c.fill = row_fill
            c.border = thin_border

    # Настройка ширины колонок
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 22
    ws.column_dimensions['C'].width = 38
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 10
    ws.column_dimensions['G'].width = 16
    ws.column_dimensions['H'].width = 18

    # Внедрение Диаграммы Ганта прямо рядом с реестром на ту же страницу (начиная с колонки J)
    chart = BarChart()
    chart.type = "bar"
    chart.dir = "bar"
    chart.style = 13
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = "Диаграмма Ганта проекта (по неделям)"
    chart.x_axis.title = "Шкала времени (недели)"
    chart.height = max(12, len(schedule_data) * 0.75)
    chart.width = 18

    data = Reference(ws, min_col=7, min_row=1, max_col=8, max_row=len(schedule_data) + 1)
    cats = Reference(ws, min_col=3, min_row=2, max_row=len(schedule_data) + 1)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    # Настройка синей гаммы для полос Ганта в Excel
    if len(chart.series) > 1:
        # Прозрачный сдвиг
        chart.series[0].graphicalProperties.solidFill = "FFFFFF"
        chart.series[0].graphicalProperties.line.solidFill = "FFFFFF"
        # Насыщенная синяя полоса выполнения (Navy Blue)
        chart.series[1].graphicalProperties.solidFill = "1F4E78"
        chart.series[1].graphicalProperties.line.solidFill = "1F4E78"

    # Размещаем график на том же листе в ячейку J1
    ws.add_chart(chart, "J1")

    filename = "SMS_Project_Weekly_Gantt_OnePage.xlsx"
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
                
                week_num = current_date.isocalendar()[1]
                week_start_monday = current_date - datetime.timedelta(days=current_date.weekday())
                
                schedule_data.append({
                    '№': idx + 1,
                    'Фаза проекта': row['Phase'],
                    'DESCRIPTION': f"{idx + 1}. {row['DESCRIPTION']}",
                    'Start': current_date,
                    'Finish': finish_date,
                    'Week_Num': week_num,
                    'Week_Label': f"W{week_num} ({week_start_monday.strftime('%d.%m')})"
                })
                current_date = get_next_business_day(finish_date, holiday_dates)

            df_sched = pd.DataFrame(schedule_data)

            st.markdown("---")
            
            # --- РАЗМЕЩЕНИЕ НА ОДНОЙ СТРАНИЦЕ (2 КОЛОНКИ ПАРАЛЛЕЛЬНО) ---
            col_left, col_right = st.columns([1, 1], gap="medium")

            # КОЛОНКА 1: Реестр задач
            with col_left:
                st.markdown("### 📋 Реестр задач")
                df_display = df_sched[['№', 'Фаза проекта', 'DESCRIPTION', 'Start', 'Finish', 'Week_Label']].copy()
                df_display['Start'] = df_display['Start'].apply(lambda x: x.strftime('%d.%m.%Y'))
                df_display['Finish'] = df_display['Finish'].apply(lambda x: x.strftime('%d.%m.%Y'))
                df_display.columns = ['№', 'Фаза', 'Описание работы', 'Старт', 'Финиш', 'Неделя']
                
                st.dataframe(
                    df_display,
                    use_container_width=True,
                    height=max(450, len(schedule_data) * 28),
                    hide_index=True
                )

            # КОЛОНКА 2: График Ганта в синей гамме
            with col_right:
                st.markdown("### 📈 График Ганта (Недели)")
                
                # Синяя цветовая палитра для фаз
                blue_palette = ["#1F4E78", "#2F5597", "#41719C", "#5B9BD5", "#8EA9DB"]

                fig = px.timeline(
                    df_sched,
                    x_start="Start",
                    x_end="Finish",
                    y="DESCRIPTION",
                    color="Фаза проекта",
                    hover_data=["№", "Start", "Finish", "Week_Label"],
                    color_discrete_sequence=blue_palette
                )
                
                fig.update_yaxes(autorange="reversed", title="")
                
                # Недельный шаг по оси X (M1 = 1 неделя)
                fig.update_xaxes(
                    title="Шкала времени (недели)",
                    dtick="M1",  
                    tickformat="%d.%m\n(W%V)",
                    showgrid=True,
                    gridcolor="#E5E8EB"
                )
                
                fig.update_layout(
                    height=max(450, len(schedule_data) * 28),
                    margin=dict(l=10, r=10, t=10, b=10),
                    plot_bgcolor="#FFFFFF",
                    legend=dict(
                        orientation="h",
                        yanchor="bottom",
                        y=1.02,
                        xanchor="right",
                        x=1,
                        title_text=""
                    ),
                    font=dict(size=11)
                )

                st.plotly_chart(fig, use_container_width=True)

            # --- ВЫГРУЗКА В EXCEL НА ОДИН ЛИСТ ---
            excel_file = create_excel_one_page(schedule_data, start_date)
            
            st.markdown("---")
            with open(excel_file, "rb") as f:
                st.download_button(
                    label="📥 СКАЧАТЬ ОДНОСТРАНИЧНЫЙ EXCEL (.XLSX) С РЕЕСТРОМ И СИНЕМ ГАНТОМ",
                    data=f,
                    file_name="SMS_Weekly_Gantt_OnePage_Blue.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.error("Задачи не найдены в исходных листах.")

