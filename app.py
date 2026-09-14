import streamlit as st
import pandas as pd
import datetime
import holidays
import plotly.express as px
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side

st.set_page_config(page_title="SMS Schedule Agent — Daily Compact Gantt", layout="wide")

st.title("📊 ИИ-Агент: Компактный SMS График (Ежедневная Диаграмма Ганта)")

@st.cache_data
def get_rf_holidays():
    current_year = datetime.datetime.now().year
    ru_holidays = holidays.RU(years=[current_year - 2, current_year - 1, current_year, current_year + 1, current_year + 2])
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

def create_excel_with_compact_gantt(schedule_data, project_start_date):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SMS График Ганта"
    ws.views.sheetView[0].showGridLines = True

    headers = ["№", "Фаза проекта", "Описание работы", "Дата начала", "Дата финиша", "Смещение (дней)", "Длительность (дней)"]
    ws.append(headers)

    header_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    header_font = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )
    
    for col_idx in range(1, len(headers) + 1):
        cell = ws.cell(row=1, column=col_idx)
        cell.fill = header_fill
        cell.font = header_font
        cell.alignment = Alignment(horizontal="center", vertical="center")

    for row_idx, item in enumerate(schedule_data, start=2):
        offset_days = (item['Start'] - project_start_date).days
        duration_days = max(1, (item['Finish'] - item['Start']).days + 1)

        ws.append([
            item['№'], item['Фаза проекта'], item['DESCRIPTION'],
            item['Start'].strftime("%d.%m.%Y"), item['Finish'].strftime("%d.%m.%Y"),
            offset_days, duration_days
        ])

        row_fill = PatternFill(start_color="F2F7FA" if row_idx % 2 == 0 else "FFFFFF", fill_type="solid")
        for col_i in range(1, len(headers) + 1):
            c = ws.cell(row=row_idx, column=col_i)
            c.fill = row_fill
            c.border = thin_border

    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 24
    ws.column_dimensions['C'].width = 40
    ws.column_dimensions['D'].width = 12
    ws.column_dimensions['E'].width = 12
    ws.column_dimensions['F'].width = 15
    ws.column_dimensions['G'].width = 15

    chart = BarChart()
    chart.type = "bar"
    chart.dir = "bar"
    chart.style = 13
    chart.grouping = "stacked"
    chart.overlap = 100
    chart.title = f"Ежедневная Диаграмма Ганта (Старт: {project_start_date.strftime('%d.%m.%Y')})"
    chart.x_axis.title = "Дни"
    chart.height = max(10, len(schedule_data) * 0.6)
    chart.width = 16

    data = Reference(ws, min_col=6, min_row=1, max_col=7, max_row=len(schedule_data) + 1)
    cats = Reference(ws, min_col=3, min_row=2, max_row=len(schedule_data) + 1)

    chart.add_data(data, titles_from_data=True)
    chart.set_categories(cats)

    if len(chart.series) > 1:
        chart.series[0].graphicalProperties.solidFill = "FFFFFF"
        chart.series[0].graphicalProperties.line.solidFill = "FFFFFF"
        chart.series[1].graphicalProperties.solidFill = "1F4E78"
        chart.series[1].graphicalProperties.line.solidFill = "1F4E78"

    ws.add_chart(chart, "I1")

    filename = "SMS_Daily_Compact_Gantt.xlsx"
    wb.save(filename)
    return filename

# --- ИНТЕРФЕЙС STREAMLIT ---
uploaded_file = st.file_uploader("Загрузите мастер-файл проекта (.xlsx)", type=["xlsx"])
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
                # Нахождение следующего рабочего дня (без пропусков внутри выполняемой задачи)
                current_date = get_next_business_day(current_date, holiday_dates)
                finish_date = current_date + datetime.timedelta(days=1)
                
                schedule_data.append({
                    '№': idx + 1,
                    'Фаза проекта': row['Phase'],
                    'DESCRIPTION': f"{idx + 1}. {row['DESCRIPTION']}",
                    'Start': current_date,
                    'Finish': finish_date
                })
                current_date = get_next_business_day(finish_date, holiday_dates)

            df_sched = pd.DataFrame(schedule_data)

            st.markdown("---")
            st.markdown("### 📈 Ежедневная Компактная Диаграмма Ганта")

            blue_palette = ["#1F4E78", "#2F5597", "#41719C", "#5B9BD5", "#8EA9DB"]

            fig = px.timeline(
                df_sched,
                x_start="Start",
                x_end="Finish",
                y="DESCRIPTION",
                color="Фаза проекта",
                hover_data=["№", "Start", "Finish"],
                color_discrete_sequence=blue_palette
            )
            
            fig.update_yaxes(autorange="reversed", title="")
            
            # Настройки оси X: ежедневный шаг, категориальный режим без вывода пустых ней (rangebreaks)
            fig.update_xaxes(
                title="Календарь (Дни)",
                dtick="D1",
                tickformat="%d.%m",
                showgrid=True,
                gridcolor="#E5E8EB",
                rangebreaks=[
                    dict(bounds=["sat", "mon"]) # Скрытие выходных (пустых) дней
                ]
            )
            
            # Компактная высотная привязка без растягивания ячеек
            fig.update_layout(
                height=max(400, len(schedule_data) * 22),
                margin=dict(l=5, r=5, t=10, b=10),
                plot_bgcolor="#FFFFFF",
                legend=dict(
                    orientation="h",
                    yanchor="bottom",
                    y=1.01,
                    xanchor="right",
                    x=1,
                    title_text=""
                ),
                font=dict(size=11)
            )

            # Отображение только диаграммы Ганта
            st.plotly_chart(fig, use_container_width=True)

            # Кнопка скачивания
            excel_file = create_excel_with_compact_gantt(schedule_data, start_date)
            
            st.markdown("---")
            with open(excel_file, "rb") as f:
                st.download_button(
                    label="📥 СКАЧАТЬ ЕЖЕДНЕВНЫЙ ГРАФИК (EXCEL .XLSX)",
                    data=f,
                    file_name="SMS_Daily_Compact_Gantt_Blue.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.error("Задачи не найдены.")
