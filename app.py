# -*- coding: utf-8 -*-
import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from datetime import datetime, timedelta
import plotly.express as px
import plotly.graph_objects as go
import io

# ============================================================
# 1. КОНСТАНТЫ И НАСТРОЙКИ
# ============================================================
SHEET_MILESTONES = "START PROJECT TOGF-ENG-007-02"
SHEET_PHASES = [
    "PMSPR TOGF-ENG-008-06 Phase2",
    "PMSPR TOGF-ENG-008-06 Phase 3",
    "PMSPR TOGF-ENG-008-06 Phase4;5",
]
DURATION_COLUMNS = ["AM", "AN", "AQ", "AP"]
MILESTONE_ROWS = range(30, 36)
MILESTONE_NAME_COL = "B"
MILESTONE_DATE_COL = "C"

RU_HOLIDAYS = {
    2024: ["2024-01-01","2024-01-02","2024-01-03","2024-01-04","2024-01-05","2024-01-06","2024-01-07","2024-01-08","2024-02-23","2024-03-08","2024-04-29","2024-04-30","2024-05-01","2024-05-09","2024-05-10","2024-06-12","2024-11-04","2024-12-30","2024-12-31"],
    2025: ["2025-01-01","2025-01-02","2025-01-03","2025-01-06","2025-01-07","2025-01-08","2025-02-24","2025-03-10","2025-05-01","2025-05-02","2025-05-08","2025-05-09","2025-06-12","2025-06-13","2025-11-03","2025-11-04","2025-12-31"],
    2026: ["2026-01-01","2026-01-02","2026-01-05","2026-01-06","2026-01-07","2026-01-08","2026-02-23","2026-03-09","2026-05-01","2026-05-11","2026-06-12","2026-11-04"],
}

# ============================================================
# 2. ФУНКЦИИ ОБРАБОТКИ ДАННЫХ
# ============================================================
def build_holiday_set(years):
    holidays = set()
    for y in years:
        for d in RU_HOLIDAYS.get(y, []):
            holidays.add(datetime.strptime(d, "%Y-%m-%d").date())
    return holidays

def is_working_day(d, holidays):
    if d.weekday() >= 5: return False
    if d in holidays: return False
    return True

def add_working_days(start_date, working_days, holidays):
    current = start_date
    added = 0
    while added < working_days:
        current += timedelta(days=1)
        if is_working_day(current, holidays):
            added += 1
    return current

def read_milestones(wb):
    ws = wb[SHEET_MILESTONES]
    milestones = []
    for row in MILESTONE_ROWS:
        name = ws[f"{MILESTONE_NAME_COL}{row}"].value
        date_val = ws[f"{MILESTONE_DATE_COL}{row}"].value
        if name and pd.notna(date_val):
            # Преобразуем дату к datetime
            if isinstance(date_val, datetime):
                m_date = date_val
            elif isinstance(date_val, str):
                try:
                    m_date = datetime.strptime(date_val, "%Y-%m-%d")
                except:
                    continue
            else:
                try:
                    m_date = pd.to_datetime(date_val)
                except:
                    continue
            milestones.append({"name": str(name).strip(), "date": m_date})
    return milestones

def count_filled_cells(ws, row, columns):
    filled = 0
    for col_letter in columns:
        cell = ws[f"{col_letter}{row}"]
        fill = cell.fill
        if fill is not None and fill.fill_type not in (None, "none"):
            fg = fill.fgColor
            if fg is not None and fg.rgb not in (None, "00000000", "FFFFFFFF"):
                filled += 1
    return filled

def find_description_column(ws):
    for row in range(1, 21):
        for col in range(1, 40):
            val = ws.cell(row=row, column=col).value
            if val and str(val).strip().upper() == "DESCRIPTION":
                return row, col
    return None, None

def read_phase_tasks(wb, sheet_name):
    ws = wb[sheet_name]
    header_row, desc_col = find_description_column(ws)
    if header_row is None:
        return []
    tasks = []
    for row in range(header_row + 1, ws.max_row + 1):
        desc = ws.cell(row=row, column=desc_col).value
        if not desc or not str(desc).strip():
            continue
        duration = count_filled_cells(ws, row, DURATION_COLUMNS)
        if duration == 0:
            continue
        tasks.append({
            "phase": sheet_name,
            "description": str(desc).strip(),
            "duration_weeks": duration,
        })
    return tasks

def build_schedule(tasks, milestones, start_date):
    years = set()
    for m in milestones:
        if isinstance(m["date"], datetime):
            years.add(m["date"].year)
    if not years:
        years = {start_date.year, start_date.year + 1}
    holidays = build_holiday_set(list(years))

    current = start_date
    for t in tasks:
        wd = t["duration_weeks"] * 5
        t_start = current
        while not is_working_day(t_start, holidays):
            t_start += timedelta(days=1)
        
        t_end = add_working_days(t_start, wd - 1, holidays) if wd > 0 else t_start
        
        t["start"] = t_start
        t["end"] = t_end
        t["duration_days"] = wd
        current = t_end + timedelta(days=1)
    return tasks

def export_to_excel(tasks, milestones):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SMS P25077"

    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    milestone_font = Font(bold=True, color="C00000", size=10)
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(border_style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    ws["A1"] = "SMS график запуска в серийное производство P25077"
    ws["A1"].font = Font(bold=True, size=14, color="1F4E78")
    ws.merge_cells("A1:G1")

    headers = ["№", "Фаза", "Описание действия", "Начало", "Окончание", "Длит. (нед.)", "Длит. (раб. дн.)"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=3, column=i, value=h)
        c.font = header_font; c.fill = header_fill; c.alignment = center; c.border = border

    row_idx = 4
    for i, t in enumerate(tasks, start=1):
        ws.cell(row=row_idx, column=1, value=i).alignment = center
        ws.cell(row=row_idx, column=2, value=t["phase"]).alignment = left
        ws.cell(row=row_idx, column=3, value=t["description"]).alignment = left
        ws.cell(row=row_idx, column=4, value=t["start"].strftime("%d.%m.%Y")).alignment = center
        ws.cell(row=row_idx, column=5, value=t["end"].strftime("%d.%m.%Y")).alignment = center
        ws.cell(row=row_idx, column=6, value=t["duration_weeks"]).alignment = center
        ws.cell(row=row_idx, column=7, value=t["duration_days"]).alignment = center
        for col in range(1, 8):
            ws.cell(row=row_idx, column=col).border = border
        row_idx += 1

    for col in ["A", "B", "C", "D", "E", "F", "G"]:
        ws.column_dimensions[col].width = 15
    ws.column_dimensions["C"].width = 50

    # Встроенная диаграмма Ганта в Excel
    gantt_start_col = 9
    gantt_start_row = 3
    min_date = min(t["start"] for t in tasks) if tasks else datetime.today().date()
    max_date = max(t["end"] for t in tasks) if tasks else datetime.today().date()
    total_days = (max_date - min_date).days + 1
    step = max(1, total_days // 100)

    ws.cell(row=gantt_start_row - 1, column=gantt_start_col, value="Диаграмма Ганта").font = Font(bold=True, size=12)

    for i, t in enumerate(tasks, start=1):
        r = gantt_start_row + i
        ws.cell(row=r, column=gantt_start_col - 1, value=i).alignment = center
        start_offset = (t["start"] - min_date).days // step
        end_offset = (t["end"] - min_date).days // step
        for offset in range(start_offset, end_offset + 1):
            cell = ws.cell(row=r, column=gantt_start_col + offset)
            cell.fill = PatternFill("solid", fgColor="4472C4")
            cell.border = border

    for m in milestones:
        m_date = m["date"].date() if isinstance(m["date"], datetime) else m["date"]
        offset = (m_date - min_date).days // step
        col = gantt_start_col + offset
        if col >= gantt_start_col:
            label_cell = ws.cell(row=gantt_start_row - 1, column=col, value=m["name"])
            label_cell.font = milestone_font
            label_cell.alignment = Alignment(horizontal="center", text_rotation=90)
            for r in range(gantt_start_row, gantt_start_row + len(tasks) + 1):
                cell = ws.cell(row=r, column=col)
                if cell.fill is None or cell.fill.fill_type in (None, "none"):
                    cell.fill = PatternFill("solid", fgColor="FFC000")

    stream = io.BytesIO()
    wb.save(stream)
    stream.seek(0)
    return stream

# ============================================================
# 3. ИНТЕРФЕЙС STREAMLIT
# ============================================================
st.set_page_config(page_title="Генератор SMS Графика P25077", layout="wide")
st.title("🏭 Генератор SMS Графика Запуска P25077")
st.markdown("Загрузите файл **PLANT_MASTER_SCHEDULE P25077.xlsx**. Система автоматически считает цветные ячейки (AM, AN, AQ, AP) как недели длительности и построит график с учетом праздников РФ.")

uploaded_file = st.file_uploader("Выберите Excel-файл", type=["xlsx", "xlsm"])

if uploaded_file is not None:
    if st.button("🚀 Сформировать график", type="primary"):
        with st.spinner("🤖 ИИ-агент анализирует файл, считает цветные ячейки и строит календарь..."):
            try:
                wb = openpyxl.load_workbook(uploaded_file, data_only=True)
                
                milestones = read_milestones(wb)
                all_tasks = []
                for sheet in SHEET_PHASES:
                    if sheet in wb.sheetnames:
                        tasks = read_phase_tasks(wb, sheet)
                        all_tasks.extend(tasks)
                
                if not all_tasks:
                    st.error("❌ Задачи не найдены. Проверьте наличие вкладки 'DESCRIPTION' и цветную заливку в столбцах AM, AN, AQ, AP.")
                else:
                    start_date = datetime.today().date()
                    for m in milestones:
                        if isinstance(m["date"], datetime):
                            start_date = m["date"].date()
                            break
                    
                    tasks = build_schedule(all_tasks, milestones, start_date)
                    df_tasks = pd.DataFrame(tasks)
                    
                    # ИСПРАВЛЕНИЕ: Явно преобразуем к datetime перед использованием .dt
                    df_tasks["start"] = pd.to_datetime(df_tasks["start"], errors='coerce')
                    df_tasks["end"] = pd.to_datetime(df_tasks["end"], errors='coerce')
                    
                    # Теперь безопасно используем .dt.strftime
                    df_tasks["start_fmt"] = df_tasks["start"].dt.strftime("%d.%m.%Y")
                    df_tasks["end_fmt"] = df_tasks["end"].dt.strftime("%d.%m.%Y")
                    
                    st.success(f"✅ Обработка завершена! Найдено задач: {len(df_tasks)}, вех: {len(milestones)}")
                    
                    col1, col2 = st.columns([1, 2])
                    
                    with col1:
                        st.subheader("📋 Таблица SMS")
                        display_df = df_tasks[["phase", "description", "start_fmt", "end_fmt", "duration_weeks", "duration_days"]].copy()
                        display_df.columns = ["Фаза", "Описание", "Начало", "Окончание", "Длит. (нед)", "Длит. (дн)"]
                        st.dataframe(display_df, use_container_width=True, hide_index=True)
                        
                        if milestones:
                            st.subheader(" Вехи проекта")
                            df_m = pd.DataFrame(milestones)
                            df_m["date_fmt"] = pd.to_datetime(df_m["date"]).dt.strftime("%d.%m.%Y")
                            st.dataframe(df_m[["name", "date_fmt"]].rename(columns={"name": "Веха", "date_fmt": "Дата"}), use_container_width=True, hide_index=True)

                        # Кнопки скачивания
                        excel_stream = export_to_excel(tasks, milestones)
                        st.download_button(
                            label="📥 Скачать профессиональный Excel с Гантом",
                            data=excel_stream,
                            file_name="SMS_P25077_Result.xlsx",
                            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                            use_container_width=True
                        )
                        
                        csv = display_df.to_csv(index=False, sep=";").encode('utf-8-sig')
                        st.download_button(
                            label=" Скачать таблицу CSV",
                            data=csv,
                            file_name="SMS_Schedule.csv",
                            mime="text/csv",
                            use_container_width=True
                        )

                    with col2:
                        st.subheader("📊 Интерактивная диаграмма Ганта")
                        fig = px.timeline(
                            df_tasks,
                            x_start="start",
                            x_end="end",
                            y="description",
                            color="phase",
                            title="График запуска в серийное производство",
                            hover_data=["duration_weeks", "duration_days"]
                        )
                        fig.update_yaxes(autorange="reversed", title="Задачи")
                        
                        # Добавляем вехи на график
                        for m in milestones:
                            fig.add_vline(
                                x=m["date"],
                                line_dash="dash",
                                line_color="red",
                                line_width=2,
                                annotation_text=f"🚩 {m['name']}",
                                annotation_position="top right"
                            )
                        
                        fig.update_layout(
                            xaxis_title="Календарный план (с учетом выходных и праздников РФ)",
                            height=max(400, len(df_tasks) * 30),
                            xaxis=dict(tickformat="%d.%m.%Y"),
                            legend_title="Фаза"
                        )
                        st.plotly_chart(fig, use_container_width=True)
                        
            except Exception as e:
                st.error(f"❌ Ошибка обработки файла: {str(e)}")
                st.exception(e)
                        
            except Exception as e:
                st.error(f"❌ Ошибка обработки файла: {str(e)}")
                st.exception(e) # Покажет детали ошибки для отладки
