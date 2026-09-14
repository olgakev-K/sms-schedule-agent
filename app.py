
import io
import datetime
import pandas as pd
import numpy as np
import plotly.express as px
import plotly.graph_objects as go
import streamlit as st
import holidays
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# ---------------------------------------------------------
# Конфигурация страницы Streamlit
# ---------------------------------------------------------
st.set_page_config(
    page_title="SMS Master Schedule AI Agent",
    page_icon="📅",
    layout="wide"
)

st.title("🤖 ИИ-Агент: График запуска проекта в серийное производство (SMS)")
st.caption("Автоматический расчёт и визуализация на основе производственного календаря РФ")

# ---------------------------------------------------------
# Строгие константы согласно ТЗ
# ---------------------------------------------------------
IGNORED_COLUMNS = [
    "PLANNED START DATE",
    "PLANNED START WEEK (automatic)",
    "PLANNED END DATE (automatic)",
    "PLANNED END WEEK (automatic)",
    "ACTUAL END DATE STATUS (automatic)"
]

PHASE_SHEETS = [
    "PMSPR TOGF-ENG-008-06 Phase2",
    "PMSPR TOGF-ENG-008-06 Phase 3",
    "PMSPR TOGF-ENG-008-06 Phase4;5"
]

MILESTONE_SHEET = "START PROJECT TOGF-ENG-007-02"

# ---------------------------------------------------------
# Производственный календарь РФ (Holidays RU)
# ---------------------------------------------------------
@st.cache_data
def get_ru_holidays(start_year=2024, end_year=2030):
    """Получить список официальных государственных праздников и переносов РФ."""
    years = list(range(start_year, end_year + 1))
    return holidays.RU(years=years)

def add_working_days(start_date, days, ru_holidays):
    """
    Прибавить рабочие дни с учётом обычных выходных (СБ, ВС) 
    и государственных праздников РФ.
    """
    if pd.isna(start_date):
        return pd.NaT
    
    curr = pd.to_datetime(start_date).date()
    added = 0
    target_days = max(1, int(days)) if not pd.isna(days) else 1
    
    while added < target_days - 1:
        curr += datetime.timedelta(days=1)
        # 0 = Понедельник, 4 = Пятница, 5 = Суббота, 6 = Воскресенье
        if curr.weekday() < 5 and curr not in ru_holidays:
            added += 1
            
    return curr

# ---------------------------------------------------------
# Движок обработки Excel-файла
# ---------------------------------------------------------
def process_excel_schedule(file_bytes):
    """
    Парсинг Excel, исключение системных колонок, 
    извлечение вех и расчёт календарного графика.
    """
    xls = pd.ExcelFile(file_bytes)
    ru_holidays = get_ru_holidays(2024, 2030)
    
    # 1. Извлечение даты старта из вех проекта (Условие 1)
    project_start_date = datetime.date.today()
    if MILESTONE_SHEET in xls.sheet_names:
        milestones_df = pd.read_excel(xls, sheet_name=MILESTONE_SHEET)
        for col in milestones_df.columns:
            # Ищем первую валидную дату в листе вех
            date_series = pd.to_datetime(milestones_df[col], errors='coerce').dropna()
            if not date_series.empty:
                project_start_date = date_series.iloc[0].date()
                break

    # 2. Парсинг задач с указанных фазовых листов (Условие 2 и 3)
    tasks_list = []
    current_date = project_start_date

    for sheet_name in PHASE_SHEETS:
        if sheet_name not in xls.sheet_names:
            continue
        
        df = pd.read_excel(xls, sheet_name=sheet_name)
        
        # Полное ИГНОРИРОВАНИЕ указанных колонок (Условие 3)
        valid_cols = [c for c in df.columns if str(c).strip() not in IGNORED_COLUMNS]
        df = df[valid_cols]
        
        # Поиск колонки DESCRIPTION
        desc_col = None
        for c in df.columns:
            if "DESCRIPTION" in str(c).upper():
                desc_col = c
                break
        
        if desc_col is None and len(df.columns) > 0:
            desc_col = df.columns[0]
            
        # Поиск колонки длительности
        dur_col = None
        for c in df.columns:
            if any(term in str(c).upper() for term in ["DURATION", "ДЛИТЕЛЬНОСТЬ", "DAYS"]):
                dur_col = c
                break

        for idx, row in df.iterrows():
            desc = str(row[desc_col]).strip() if desc_col and pd.notna(row[desc_col]) else ""
            if not desc or desc.lower() in ["nan", "none", ""]:
                continue
            
            # Длительность по умолчанию, если не указана
            duration = 5
            if dur_col and pd.notna(row[dur_col]):
                try:
                    duration = max(1, int(float(row[dur_col])))
                except ValueError:
                    duration = 5
            
            # Корректировка даты старта задачи (если выпадает на праздник/выходной)
            while current_date.weekday() >= 5 or current_date in ru_holidays:
                current_date += datetime.timedelta(days=1)
                
            start_dt = current_date
            end_dt = add_working_days(start_dt, duration, ru_holidays)
            
            tasks_list.append({
                "Phase": sheet_name,
                "Description": desc,
                "Duration (Days)": duration,
                "Start Date": start_dt,
                "End Date": end_dt
            })
            
            # Дата старта следующей задачи
            next_start = end_dt + datetime.timedelta(days=1)
            while next_start.weekday() >= 5 or next_start in ru_holidays:
                next_start += datetime.timedelta(days=1)
            current_date = next_start

    schedule_df = pd.DataFrame(tasks_list)
    return schedule_df, project_start_date

# ---------------------------------------------------------
# Генератор Excel с диаграммой Ганта прямо на листе
# ---------------------------------------------------------
def generate_excel_with_gantt(schedule_df):
    """
    Формирует XLS файл с таблицей задач и ячеистой
    диаграммой Ганта (цветовая заливка по неделям).
    """
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SMS_Schedule"
    ws.views.sheetView[0].showGridLines = True

    # Цветовая палитра и стили
    HEADER_FILL = PatternFill(start_color="1F497D", end_color="1F497D", fill_type="solid")
    WEEK_HEADER_FILL = PatternFill(start_color="366092", end_color="366092", fill_type="solid")
    GANTT_FILL = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid")
    ALT_ROW_FILL = PatternFill(start_color="F2F5F9", end_color="F2F5F9", fill_type="solid")
    
    HEADER_FONT = Font(name="Calibri", size=11, bold=True, color="FFFFFF")
    REG_FONT = Font(name="Calibri", size=10)
    
    THIN_BORDER = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Заголовок документа
    ws.cell(row=1, column=1, value="ГРАФИК ЗАПУСКА ПРОЕКТА В СЕРИЙНОЕ ПРОИЗВОДСТВО (SMS)").font = Font(name="Calibri", size=14, bold=True, color="1F497D")
    ws.cell(row=2, column=1, value=f"Сформировано: {datetime.date.today().strftime('%d.%m.%Y')} | Производственный календарь РФ").font = Font(name="Calibri", size=10, italic=True)

    # Табличные заголовки
    table_headers = ["№", "Фаза (Phase)", "Описание работ (Description)", "Длит. (раб. дн.)", "Дата начала", "Дата окончания"]
    start_row = 4
    
    for col_idx, header in enumerate(table_headers, start=1):
        cell = ws.cell(row=start_row, column=col_idx, value=header)
        cell.fill = HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Таймлайн по неделям для диаграммы Ганта
    min_date = schedule_df["Start Date"].min()
    max_date = schedule_df["End Date"].max()
    
    weeks = []
    curr = min_date
    while curr <= max_date + datetime.timedelta(days=7):
        w_num = curr.isocalendar()[1]
        w_label = f"W{w_num}\n({curr.strftime('%d.%m')})"
        if not any(w[0] == w_num for w in weeks):
            weeks.append((w_num, w_label, curr))
        curr += datetime.timedelta(days=7)

    gantt_start_col = len(table_headers) + 1
    for i, (w_num, w_label, w_date) in enumerate(weeks):
        col_idx = gantt_start_col + i
        cell = ws.cell(row=start_row, column=col_idx, value=w_label)
        cell.fill = WEEK_HEADER_FILL
        cell.font = HEADER_FONT
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(col_idx)].width = 9

    # Заполнение строк и построение визуальной полосы Ганта
    for row_idx, row in schedule_df.iterrows():
        r = start_row + 1 + row_idx
        ws.cell(row=r, column=1, value=row_idx + 1).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=row["Phase"])
        ws.cell(row=r, column=3, value=row["Description"])
        ws.cell(row=r, column=4, value=row["Duration (Days)"]).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=5, value=row["Start Date"].strftime("%d.%m.%Y")).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=6, value=row["End Date"].strftime("%d.%m.%Y")).alignment = Alignment(horizontal="center")

        for c in range(1, len(table_headers) + 1):
            cell = ws.cell(row=r, column=c)
            cell.font = REG_FONT
            cell.border = THIN_BORDER
            if row_idx % 2 == 1:
                cell.fill = ALT_ROW_FILL

        # Заливка ячеек Диаграммы Ганта
        task_start = row["Start Date"]
        task_end = row["End Date"]

        for i, (w_num, w_label, w_date) in enumerate(weeks):
            col_idx = gantt_start_col + i
            cell = ws.cell(row=r, column=col_idx)
            cell.border = THIN_BORDER
            
            w_start = w_date - datetime.timedelta(days=w_date.weekday())
            w_end = w_start + datetime.timedelta(days=6)
            
            # Если задача пересекает данную неделю
            if not (task_end < w_start or task_start > w_end):
                cell.fill = GANTT_FILL

    # Настройка ширины основных столбцов
    ws.column_dimensions['A'].width = 5
    ws.column_dimensions['B'].width = 28
    ws.column_dimensions['C'].width = 45
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 14
    ws.column_dimensions['F'].width = 14

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# ---------------------------------------------------------
# Streamlit Интерфейс
# ---------------------------------------------------------
uploaded_file = st.file_uploader(
    "Загрузите исходный файл Excel (PLANT_MASTER_SCHEDULE P25077.xlsx)", 
    type=["xlsx"]
)

if uploaded_file is not None:
    with st.spinner("Извлечение вех, фильтрация данных и расчёт по календарю РФ..."):
        schedule_df, start_dt = process_excel_schedule(uploaded_file)
    
    if not schedule_df.empty:
        st.success(f"График успешно сформирован! Дата старта проекта (из вех): **{start_dt.strftime('%d.%m.%Y')}**")
        
        # Метрики
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Всего этапов", len(schedule_df))
        col2.metric("Дата старта", schedule_df["Start Date"].min().strftime("%d.%m.%Y"))
        col3.metric("Дата завершения", schedule_df["End Date"].max().strftime("%d.%m.%Y"))
        col4.metric("Календарных дней", (schedule_df["End Date"].max() - schedule_df["Start Date"].min()).days)

        st.markdown("---")
        st.subheader("📊 Интерактивная Диаграмма Ганта (Plotly)")
        
        # Plotly Gantt Chart
        fig = px.timeline(
            schedule_df,
            x_start="Start Date",
            x_end="End Date",
            y="Description",
            color="Phase",
            title="График запуска в серийное производство (SMS)",
            hover_data=["Duration (Days)", "Start Date", "End Date"]
        )
        fig.update_yaxes(autorange="reversed")
        fig.update_layout(
            height=300 + len(schedule_df) * 25,
            xaxis_title="Временная шкала",
            yaxis_title="Задачи",
            legend_title="Фаза проекта"
        )
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("📋 Детализированный график работ")
        display_df = schedule_df.copy()
        display_df["Start Date"] = display_df["Start Date"].apply(lambda x: x.strftime("%d.%m.%Y"))
        display_df["End Date"] = display_df["End Date"].apply(lambda x: x.strftime("%d.%m.%Y"))
        st.dataframe(display_df, use_container_width=True)

        # Выгрузка Excel файла с диаграммой Ганта внутри
        excel_data = generate_excel_with_gantt(schedule_df)
        st.download_button(
            label="📥 Скачать Excel с Диаграммой Ганта",
            data=excel_data,
            file_name=f"SMS_Schedule_P25077_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Не удалось найти подходящие задачи на указанных листах фаз.")
else:
    st.info("Пожалуйста, загрузите Excel-файл для автоматического построения SMS-графика.")
