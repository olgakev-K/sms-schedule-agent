import io
import datetime
from datetime import timedelta
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import plotly.express as px
import streamlit as st
import holidays

# -----------------------------------------------------------------------------
# 1. КОНФИГУРАЦИЯ СТРАНИЦЫ
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="SMS-график проекта (TOGF-ENG)",
    page_icon="📅",
    layout="wide"
)

st.title("📅 Генератор SMS-графика проекта запуска (TOGF-ENG)")
st.caption("Формирование точного графика с учетом рабочих дней и государственных праздников РФ")

# -----------------------------------------------------------------------------
# 2. РАБОТА С ПРОИЗВОДСТВЕННЫМ КАЛЕНДАРЕМ РФ (Рабочие vs Выходные дни)
# -----------------------------------------------------------------------------
@st.cache_data
def get_ru_holidays_set(year_start: int, year_end: int):
    """Множество официальных праздников РФ (holidays RU / Consultant.ru)."""
    ru_holidays = set()
    for yr in range(year_start, year_end + 1):
        for d in holidays.RU(years=yr).keys():
            if isinstance(d, datetime.datetime):
                ru_holidays.add(d.date())
            else:
                ru_holidays.add(d)
    return ru_holidays

def is_workday(dt: datetime.date, ru_holidays: set) -> bool:
    """Проверка: является ли день рабочим (не СБ, не ВС и не праздник)."""
    if dt.weekday() in (5, 6): # 5 = Суббота, 6 = Воскресенье
        return False
    if dt in ru_holidays: # Праздник РФ
        return False
    return True

def add_workdays(start_dt: datetime.date, num_workdays: int, ru_holidays: set) -> datetime.date:
    """Прибавляет к дате строго N рабочих дней."""
    curr = start_dt
    added = 0
    # Если начальный день выпадает на выходной/праздник — сдвигаем на первый рабочий
    while not is_workday(curr, ru_holidays):
        curr += timedelta(days=1)
        
    while added < num_workdays:
        if is_workday(curr, ru_holidays):
            added += 1
            if added == num_workdays:
                break
        curr += timedelta(days=1)
    return curr

def get_next_workday(dt: datetime.date, ru_holidays: set) -> datetime.date:
    """Находит ближайший следующий рабочий день."""
    curr = dt + timedelta(days=1)
    while not is_workday(curr, ru_holidays):
        curr += timedelta(days=1)
    return curr

# -----------------------------------------------------------------------------
# 3. ТОЧНЫЙ ПАРСИНГ ФАЙЛА И РАСЧЕТ ДАТ
# -----------------------------------------------------------------------------
def process_excel_schedule(file_bytes, project_start_date: datetime.date):
    wb = openpyxl.load_workbook(filename=io.BytesIO(file_bytes), data_only=True)
    
    # --- Условие 1: Вехи из "START PROJECT TOGF-ENG-007-02" ---
    milestones = []
    milestone_sheet = "START PROJECT TOGF-ENG-007-02"
    if milestone_sheet in wb.sheetnames:
        ws_m = wb[milestone_sheet]
        for row in ws_m.iter_rows(values_only=True):
            vals = [str(cell).strip() for cell in row if cell is not None and str(cell).strip() != ""]
            if vals:
                milestones.append(" | ".join(vals[:3]))
    
    # --- Условие 2: Строгий порядок вкладок Phase2 -> Phase 3 -> Phase4;5 ---
    target_sheets = [
        ("PMSPR TOGF-ENG-008-06 Phase2", "Phase 2"),
        ("PMSPR TOGF-ENG-008-06 Phase 3", "Phase 3"),
        ("PMSPR TOGF-ENG-008-06 Phase4;5", "Phase 4;5")
    ]
    
    tasks = []
    # Колонка AM = 39-я колонка
    start_matrix_col = 39 
    
    for sheet_name, phase_label in target_sheets:
        if sheet_name not in wb.sheetnames:
            continue
            
        ws = wb[sheet_name]
        desc_col_idx = None
        header_row = 1
        
        # Нахождение колонки DESCRIPTION (Игнорируя автоматические даты по Условию 3)
        for r in range(1, min(20, ws.max_row + 1)):
            for c in range(1, ws.max_column + 1):
                val = str(ws.cell(row=r, column=c).value or "").strip().upper()
                if "DESCRIPTION" in val or "ОПИСАНИЕ" in val or "ДЕЙСТВИЕ" in val:
                    desc_col_idx = c
                    header_row = r
                    break
            if desc_col_idx:
                break
                
        if not desc_col_idx:
            desc_col_idx = 1
            
        for r in range(header_row + 1, ws.max_row + 1):
            desc_val = ws.cell(row=r, column=desc_col_idx).value
            if not desc_val or str(desc_val).strip() == "":
                continue
                
            task_desc = str(desc_val).strip()
            duration_weeks = 0
            max_check_col = max(start_matrix_col + 80, ws.max_column + 1)
            
            # Подсчет длительности (смежные ячейки матричной части)
            for c in range(start_matrix_col, max_check_col):
                cell = ws.cell(row=r, column=c)
                fill = cell.fill
                
                has_fill = False
                if fill and fill.fill_type and fill.fill_type != 'none':
                    fg = getattr(fill.start_color, 'rgb', None) or getattr(fill.start_color, 'theme', None)
                    if fg and str(fg) not in ['00000000', 'FFFFFFFF', '0']:
                        has_fill = True
                
                has_val = cell.value is not None and str(cell.value).strip() != ""
                
                if has_fill or has_val:
                    duration_weeks += 1
                else:
                    if duration_weeks > 0:
                        break
            
            if duration_weeks == 0:
                duration_weeks = 1 # Минимальная длительность 1 неделя
                
            tasks.append({
                "phase": phase_label,
                "description": task_desc,
                "duration_weeks": duration_weeks
            })

    if not tasks:
        return None, None, None

    # --- Условие 4: Учет производственного календаря РФ при расчете дат ---
    ru_holidays = get_ru_holidays_set(project_start_date.year, project_start_date.year + 4)
    
    current_start = project_start_date
    if not is_workday(current_start, ru_holidays):
        current_start = get_next_workday(current_start, ru_holidays)
        
    processed_tasks = []
    
    for idx, t in enumerate(tasks, 1):
        # 1 неделя длительности = 5 рабочих дней
        workdays_count = t["duration_weeks"] * 5
        
        # Дата окончания рассчитывается по РАБОЧИМ ДНЯМ
        task_end = add_workdays(current_start, workdays_count, ru_holidays)
        
        processed_tasks.append({
            "id": idx,
            "phase": t["phase"],
            "description": t["description"],
            "duration_weeks": t["duration_weeks"],
            "workdays": workdays_count,
            "start_date": current_start,
            "end_date": task_end,
        })
        
        # Следующая задача стартует строго в БЛИЖАЙШИЙ РАБОЧИЙ ДЕНЬ
        current_start = get_next_workday(task_end, ru_holidays)

    df_tasks = pd.DataFrame(processed_tasks)
    
    # --- Построение недель для матрицы Excel ---
    max_end_date = df_tasks["end_date"].max()
    weeks_info = []
    w_start = project_start_date
    w_num = 1
    
    while w_start <= max_end_date + timedelta(days=7):
        w_end = w_start + timedelta(days=6)
        
        # Проверка недели на праздники РФ
        has_holiday = False
        curr = w_start
        while curr <= w_end:
            if curr in ru_holidays:
                has_holiday = True
                break
            curr += timedelta(days=1)
            
        weeks_info.append({
            "week_num": w_num,
            "week_label": f"W{w_num}",
            "start_date": w_start,
            "end_date": w_end,
            "has_holiday": has_holiday
        })
        w_start = w_end + timedelta(days=1)
        w_num += 1

    return df_tasks, weeks_info, milestones

# -----------------------------------------------------------------------------
# 4. ФОРМИРОВАНИЕ EXCEL С ДИАГРАММОЙ ГАНТА ПО ЯЧЕЙКАМ
# -----------------------------------------------------------------------------
def generate_excel_report(df_tasks, weeks_info, milestones):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SMS Schedule"
    ws.views.sheetView[0].showGridLines = True
    
    font_header = Font(name="Arial", size=9, bold=True, color="FFFFFF")
    font_body = Font(name="Arial", size=9)
    font_bold = Font(name="Arial", size=9, bold=True)
    
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_gantt_bar = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_holiday = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Заголовки
    headers = ["№", "Фаза (Phase)", "Описание действия (Description)", "Длит. (нед)", "Раб. дней", "Дата начала", "Дата окончания"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    # Колонки недель
    start_week_col = len(headers) + 1
    for i, w in enumerate(weeks_info):
        col_idx = start_week_col + i
        label = f"{w['week_label']} 🎉" if w["has_holiday"] else w["week_label"]
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.font = Font(name="Arial", size=9, bold=True, color="9C0006" if w["has_holiday"] else "FFFFFF")
        cell.fill = fill_holiday if w["has_holiday"] else fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
        # Ширина столбца недели = 4
        ws.column_dimensions[get_column_letter(col_idx)].width = 4

    # Заполнение задач
    for row_idx, task in df_tasks.iterrows():
        r = row_idx + 2
        vals = [
            task["id"],
            task["phase"],
            task["description"],
            task["duration_weeks"],
            task["workdays"],
            task["start_date"].strftime("%d.%m.%Y"),
            task["end_date"].strftime("%d.%m.%Y")
        ]
        
        for c_idx, val in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c_idx, value=val)
            cell.font = font_body
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center" if c_idx in [1, 4, 5, 6, 7] else "left", vertical="center")

        # Рисование ячеек Ганта
        for w_idx, w in enumerate(weeks_info):
            c_idx = start_week_col + w_idx
            cell = ws.cell(row=r, column=c_idx)
            cell.border = thin_border
            
            if not (task["end_date"] < w["start_date"] or task["start_date"] > w["end_date"]):
                cell.fill = fill_gantt_bar
                cell.value = "█"
                cell.font = Font(name="Arial", size=9, color="2F5597")
                cell.alignment = Alignment(horizontal="center", vertical="center")

    base_widths = [6, 15, 50, 11, 11, 13, 13]
    for idx, w in enumerate(base_widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    ws.freeze_panes = "H2"

    # Лист с вехами
    if milestones:
        ws_m = wb.create_sheet(title="Milestones")
        ws_m.views.sheetView[0].showGridLines = True
        ws_m.cell(row=1, column=1, value="Вехи проекта (START PROJECT TOGF-ENG-007-02)").font = font_bold
        for idx, m_text in enumerate(milestones, 2):
            ws_m.cell(row=idx, column=1, value=m_text).font = font_body

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# -----------------------------------------------------------------------------
# 5. ИНТЕРФЕЙС И КНОПКА СКАЧИВАНИЯ
# -----------------------------------------------------------------------------
st.sidebar.header("⚙️ Входные параметры")

uploaded_file = st.sidebar.file_uploader(
    "Загрузите PLANT_MASTER_SCHEDULE P25077.xlsx", 
    type=["xlsx"]
)

start_date_input = st.sidebar.date_input(
    "Дата начала проекта",
    value=datetime.date.today(),
    format="DD.MM.YYYY"
)

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    
    with st.spinner("Агент формирует график с учетом производственного календаря РФ..."):
        df_tasks, weeks_info, milestones = process_excel_schedule(file_bytes, start_date_input)
    
    if df_tasks is not None and not df_tasks.empty:
        # Метрики
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Всего задач", len(df_tasks))
        c2.metric("Сумма рабоч. дней", df_tasks["workdays"].sum())
        c3.metric("Старт проекта", df_tasks["start_date"].min().strftime("%d.%m.%Y"))
        c4.metric("Окончание", df_tasks["end_date"].max().strftime("%d.%m.%Y"))

        st.markdown("---")

        # График Plotly
        st.subheader("📊 Интерактивный SMS-график (Гант)")
        fig = px.timeline(
            df_tasks,
            x_start="start_date",
            x_end="end_date",
            y="description",
            color="phase",
            hover_data=["duration_weeks", "workdays"],
            labels={"description": "Описание действия", "phase": "Фаза", "workdays": "Раб. дней"}
        )
        fig.update_yaxes(autorange="reversed")
        fig.update_layout(height=min(800, 150 + len(df_tasks) * 25))
        st.plotly_chart(fig, use_container_width=True)

        # Таблица
        with st.expander("📋 Детализация календарного плана (Таблица)"):
            st.dataframe(
                df_tasks[["id", "phase", "description", "duration_weeks", "workdays", "start_date", "end_date"]],
                use_container_width=True
            )

        # Скачивание Excel
        st.markdown("---")
        st.subheader("📥 Скачать SMS-график")
        
        excel_file_data = generate_excel_report(df_tasks, weeks_info, milestones)
        
        st.download_button(
            label="💾 Скачать итоговый Excel-файл с диаграммой Ганта",
            data=excel_file_data,
            file_name=f"SMS_Schedule_P25077_{start_date_input.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
    else:
        st.error("Не удалось прочитать задачи. Убедитесь, что в файле есть вкладки 'PMSPR TOGF-ENG-008-06 Phase2', 'Phase 3' или 'Phase4;5'.")
else:
    st.info("👈 Загрузите исходный Excel-файл на панели слева.")
