import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import plotly.express as px
import datetime
from datetime import timedelta
import holidays
import io

# ---------------------------------------------------------
# Конфигурация страницы Streamlit
# ---------------------------------------------------------
st.set_page_config(
    page_title="ИИ-Агент: SMS График запуска производства",
    page_icon="⚙️",
    layout="wide"
)

st.title("⚙️ ИИ-Агент: Формирование SMS-графика производства")
st.markdown("""
Автоматический расчет и построение календарного плана запуска проекта в серийное производство 
на основе файла **PLANT_MASTER_SCHEDULE P25077** и **Производственного календаря РФ**.
""")

# ---------------------------------------------------------
# Константы ТЗ
# ---------------------------------------------------------
IGNORED_COLUMNS = [
    "PLANNED START DATE",
    "PLANNED START WEEK (AUTOMATIC)",
    "PLANNED END DATE (AUTOMATIC)",
    "PLANNED END WEEK (AUTOMATIC)",
    "ACTUAL END DATE STATUS (AUTOMATIC)"
]

PHASE_SHEETS = [
    "PMSPR TOGF-ENG-008-06 Phase2",
    "PMSPR TOGF-ENG-008-06 Phase 3",
    "PMSPR TOGF-ENG-008-06 Phase4;5"
]

# ---------------------------------------------------------
# Вспомогательные функции для работы с датами и цветом
# ---------------------------------------------------------
def is_colored(cell) -> bool:
    """Проверяет наличие реальной цветовой заливки в ячейке."""
    if not cell.fill or cell.fill.fill_type in (None, "none"):
        return False
    fill = cell.fill
    if fill.start_color:
        # Проверка RGB
        if fill.start_color.rgb:
            rgb = str(fill.start_color.rgb).upper()
            if rgb not in ["00000000", "FFFFFFFF", "00FFFFFF", "FFFFFF", "000000"]:
                return True
        # Проверка Theme / Indexed
        if fill.start_color.theme is not None:
            return True
        if fill.start_color.indexed is not None and fill.start_color.indexed != 64:
            return True
    return False

@st.cache_data
def get_russian_holidays(start_year=2024, end_year=2030):
    """Возвращает множество праздников РФ."""
    ru_holidays = set()
    for yr in range(start_year, end_year + 1):
        for h_date in holidays.RU(years=yr).keys():
            ru_holidays.add(h_date)
    return ru_holidays

def get_workday_range_for_weeks(project_start: datetime.date, start_week_offset: int, num_weeks: int, ru_holidays):
    """
    Вычисляет точные даты начала и окончания работы с учетом праздников РФ.
    start_week_offset: 0 для первой недели с заливкой, 1 для второй и т.д.
    """
    # Гарантируем, что дата старта — рабочий день
    curr = project_start
    while curr.weekday() >= 5 or curr in ru_holidays:
        curr += timedelta(days=1)
        
    # Находим рабочий день начала заданного диапазона недель
    # В 1 неделе 5 рабочих дней
    start_work_day_index = start_week_offset * 5
    total_work_days = num_weeks * 5
    
    # Сдвигаем начальную дату
    task_start = curr
    added_days = 0
    while added_days < start_work_day_index:
        task_start += timedelta(days=1)
        if task_start.weekday() < 5 and task_start not in ru_holidays:
            added_days += 1
            
    # Находим дату окончания (последний рабочий день задачи)
    task_end = task_start
    added_days = 1
    while added_days < total_work_days:
        task_end += timedelta(days=1)
        if task_end.weekday() < 5 and task_end not in ru_holidays:
            added_days += 1
            
    return task_start, task_end

# ---------------------------------------------------------
# Обработка Excel
# ---------------------------------------------------------
def process_excel(uploaded_file):
    # Загружаем без data_only, чтобы корректно читать стили (PatternFill)
    wb = openpyxl.load_workbook(uploaded_file, data_only=False)
    ru_holidays = get_russian_holidays(2024, 2030)

    # 1. Точное извлечение даты вехи из вкладки START PROJECT TOGF-ENG-007-02
    start_sheet_name = None
    for sheet in wb.sheetnames:
        if "START PROJECT" in sheet.upper():
            start_sheet_name = sheet
            break

    project_start_date = None
    if start_sheet_name:
        ws_start = wb[start_sheet_name]
        for row in ws_start.iter_rows(values_only=True):
            for cell_val in row:
                if isinstance(cell_val, (datetime.datetime, datetime.date)):
                    project_start_date = cell_val if isinstance(cell_val, datetime.date) else cell_val.date()
                    break
            if project_start_date:
                break

    if not project_start_date:
        project_start_date = datetime.date.today()
        st.warning(f"Дата вехи не найдена на листе. Установлена текущая дата: {project_start_date.strftime('%d.%m.%Y')}")

    tasks = []

    # 2. Обработка фазовых вкладок
    for target_phase in PHASE_SHEETS:
        matched_sheet = None
        for sheet_name in wb.sheetnames:
            if sheet_name.strip().replace(" ", "").upper() in target_phase.strip().replace(" ", "").upper():
                matched_sheet = sheet_name
                break

        if not matched_sheet:
            continue

        ws = wb[matched_sheet]

        # Находим колонку DESCRIPTION и первую колонку таймлайна
        desc_col_idx = None
        header_row_idx = 1

        for r_idx in range(1, min(20, ws.max_row + 1)):
            for c_idx in range(1, ws.max_column + 1):
                val = ws.cell(row=r_idx, column=c_idx).value
                if val and "DESCRIPTION" in str(val).strip().upper():
                    desc_col_idx = c_idx
                    header_row_idx = r_idx
                    break
            if desc_col_idx:
                break

        if not desc_col_idx:
            desc_col_idx = 2

        # Ищем первый столбец сетки недель (обычно следует за системными колонками)
        timeline_start_col = desc_col_idx + 1
        for c_idx in range(desc_col_idx + 1, ws.max_column + 1):
            val = ws.cell(row=header_row_idx, column=c_idx).value
            if val and any(ign.lower() in str(val).lower() for ign in IGNORED_COLUMNS):
                continue
            # Находим заголовок, обозначающий неделю/дату
            timeline_start_col = c_idx
            break

        # Сканируем строки задач
        for r_idx in range(header_row_idx + 1, ws.max_row + 1):
            desc_val = ws.cell(row=r_idx, column=desc_col_idx).value
            if not desc_val or str(desc_val).strip() == "":
                continue

            desc_str = str(desc_val).strip()

            # Игнорируем заголовочные строки
            if any(ign.lower() in desc_str.lower() for ign in IGNORED_COLUMNS):
                continue

            # Определяем столбцы с заливкой для ДАННОЙ строки
            colored_col_indices = []
            for c_idx in range(timeline_start_col, ws.max_column + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                if is_colored(cell):
                    colored_col_indices.append(c_idx)

            if not colored_col_indices:
                # Если заливки нет, пропускаем или ставим 1 неделю
                continue

            # Считаем Duration weeks как количество закрашенных ячеек
            duration_weeks = len(colored_col_indices)
            
            # Определяем смещение первой закрашенной ячейки относительно начала сетки
            first_colored_col = colored_col_indices[0]
            start_week_offset = first_colored_col - timeline_start_col

            # Расчет точных дат начала и окончания по производственному календарю РФ
            task_start, task_end = get_workday_range_for_weeks(
                project_start_date, start_week_offset, duration_weeks, ru_holidays
            )

            tasks.append({
                "Phase": matched_sheet,
                "Description": desc_str,
                "Duration_Weeks": duration_weeks,
                "Start_Date": task_start,
                "End_Date": task_end,
                "Start_Week_Offset": start_week_offset
            })

    return pd.DataFrame(tasks), project_start_date

# ---------------------------------------------------------
# Экспорт в Excel с цветной Диаграммой Ганта
# ---------------------------------------------------------
def create_excel_with_gantt(df_tasks, project_start_date):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SMS Master Schedule"
    ws.views.sheetView[0].showGridLines = True

    # Стили
    HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    WEEK_FILL = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    GANTT_FILL = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
    ALT_ROW_FILL = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")
    
    HEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    REG_FONT = Font(name="Calibri", size=10)
    BORDER_THIN = Border(
        left=Side(style='thin', color='D9D9D9'), right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'), bottom=Side(style='thin', color='D9D9D9')
    )

    ws.cell(row=1, column=1, value="ГРАФИК ЗАПУСКА В СЕРИЙНОЕ ПРОИЗВОДСТВО (SMS)").font = Font(name="Calibri", size=14, bold=True, color="1F4E79")
    ws.cell(row=2, column=1, value=f"Старт проекта: {project_start_date.strftime('%d.%m.%Y')} | Производственный календарь РФ").font = Font(name="Calibri", size=10, italic=True)

    headers = ["№", "Фаза проекта", "Описание действия (DESCRIPTION)", "Длительность (нед.)", "Дата начала", "Дата окончания"]
    start_row = 4

    for c_idx, text in enumerate(headers, start=1):
        c = ws.cell(row=start_row, column=c_idx, value=text)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Максимальное кол-во недель для построения шапки Ганта
    max_weeks = int((df_tasks["Start_Week_Offset"] + df_tasks["Duration_Weeks"]).max()) if not df_tasks.empty else 20
    gantt_start_col = len(headers) + 1

    for w_idx in range(max_weeks):
        col = gantt_start_col + w_idx
        c = ws.cell(row=start_row, column=col, value=f"W{w_idx + 1}")
        c.fill = WEEK_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(col)].width = 4

    for idx, row in df_tasks.iterrows():
        r = start_row + 1 + idx
        ws.cell(row=r, column=1, value=idx + 1).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=row["Phase"])
        ws.cell(row=r, column=3, value=row["Description"])
        ws.cell(row=r, column=4, value=row["Duration_Weeks"]).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=5, value=row["Start_Date"].strftime("%d.%m.%Y")).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=6, value=row["End_Date"].strftime("%d.%m.%Y")).alignment = Alignment(horizontal="center")

        for c_idx in range(1, len(headers) + 1):
            c = ws.cell(row=r, column=c_idx)
            c.font = REG_FONT
            c.border = BORDER_THIN
            if idx % 2 == 1:
                c.fill = ALT_ROW_FILL

        # Закрашиваем ровно те недели, где стоит заливка
        start_w = int(row["Start_Week_Offset"])
        dur_w = int(row["Duration_Weeks"])
        for w in range(dur_w):
            gc = ws.cell(row=r, column=gantt_start_col + start_w + w)
            gc.fill = GANTT_FILL
            gc.border = BORDER_THIN

    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 28
    ws.column_dimensions['C'].width = 50
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 14
    ws.column_dimensions['F'].width = 14

    wb.save(output)
    output.seek(0)
    return output

# ---------------------------------------------------------
# UI Streamlit
# ---------------------------------------------------------
uploaded_file = st.file_uploader("Загрузите шаблон PLANT_MASTER_SCHEDULE P25077.xlsx", type=["xlsx"])

if uploaded_file:
    with st.spinner("Точный анализ цвета ячеек и перерасчет дат по производственному календарю РФ..."):
        df_tasks, project_start_date = process_excel(uploaded_file)

    if not df_tasks.empty:
        st.success(f"Базовая дата вехи из листа START PROJECT: **{project_start_date.strftime('%d.%m.%Y')}**")

        st.subheader("📊 Интерактивная Диаграмма Ганта")
        fig = px.timeline(
            df_tasks,
            x_start="Start_Date",
            x_end="End_Date",
            y="Description",
            color="Phase",
            title="График запуска производства",
            hover_data=["Duration_Weeks", "Start_Date", "End_Date"]
        )
        fig.update_yaxes(autorange="reversed")
        fig.update_layout(height=350 + len(df_tasks) * 22)
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("📋 Таблица задач с перерасчитанными датами")
        display_df = df_tasks.copy()
        display_df["Start_Date"] = display_df["Start_Date"].apply(lambda x: x.strftime("%d.%m.%Y"))
        display_df["End_Date"] = display_df["End_Date"].apply(lambda x: x.strftime("%d.%m.%Y"))
        st.dataframe(display_df[["Phase", "Description", "Duration_Weeks", "Start_Date", "End_Date"]], use_container_width=True)

        excel_file = create_excel_with_gantt(df_tasks, project_start_date)
        st.download_button(
            label="📥 Скачать Excel с корректной Диаграммой Ганта",
            data=excel_file,
            file_name=f"SMS_Schedule_Corrected_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.error("Не удалось извлечь задачи из фазовых листов или не найдены закрашенные ячейки.")
