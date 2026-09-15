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
st.caption("Точный расчет параллельных задач по номерам недель (W) и производственному календарю РФ")

# -----------------------------------------------------------------------------
# 2. КАЛЕНДАРЬ РФ И РАБОЧИЕ ДНИ
# -----------------------------------------------------------------------------
@st.cache_data
def get_ru_holidays_set(year_start: int, year_end: int):
    """Множество государственных праздников РФ."""
    ru_holidays = set()
    for yr in range(year_start, year_end + 1):
        for d in holidays.RU(years=yr).keys():
            if isinstance(d, datetime.datetime):
                ru_holidays.add(d.date())
            else:
                ru_holidays.add(d)
    return ru_holidays

def check_week_has_holidays(week_start: datetime.date, week_end: datetime.date, ru_holidays: set) -> bool:
    """Проверка наличия официального праздника РФ на неделе."""
    curr = week_start
    while curr <= week_end:
        if curr in ru_holidays:
            return True
        curr += timedelta(days=1)
    return False

# -----------------------------------------------------------------------------
# 3. ПАРСИНГ И МАТРИЧНЫЙ РАСЧЕТ ДАТ
# -----------------------------------------------------------------------------
def process_excel_schedule(file_bytes, project_start_date: datetime.date):
    wb = openpyxl.load_workbook(filename=io.BytesIO(file_bytes), data_only=True)
    
    # 1. Вехи проекта (START PROJECT TOGF-ENG-007-02)
    milestones = []
    milestone_sheet = "START PROJECT TOGF-ENG-007-02"
    if milestone_sheet in wb.sheetnames:
        ws_m = wb[milestone_sheet]
        for row in ws_m.iter_rows(values_only=True):
            vals = [str(cell).strip() for cell in row if cell is not None and str(cell).strip() != ""]
            if vals:
                milestones.append(" | ".join(vals[:3]))
    
    # 2. Целевые вкладки
    target_sheets = [
        ("PMSPR TOGF-ENG-008-06 Phase2", "Phase 2"),
        ("PMSPR TOGF-ENG-008-06 Phase 3", "Phase 3"),
        ("PMSPR TOGF-ENG-008-06 Phase4;5", "Phase 4;5")
    ]
    
    tasks = []
    ru_holidays = get_ru_holidays_set(project_start_date.year, project_start_date.year + 5)
    
    for sheet_name, phase_label in target_sheets:
        if sheet_name not in wb.sheetnames:
            continue
            
        ws = wb[sheet_name]
        desc_col_idx = None
        header_row_idx = None
        
        # Шаг А: Поиск строки с описанием задач (DESCRIPTION)
        for r in range(1, min(30, ws.max_row + 1)):
            for c in range(1, min(20, ws.max_column + 1)):
                val = str(ws.cell(row=r, column=c).value or "").strip().upper()
                if "DESCRIPTION" in val or "ОПИСАНИЕ" in val or "ДЕЙСТВИЕ" in val:
                    desc_col_idx = c
                    header_row_idx = r
                    break
            if desc_col_idx:
                break
                
        if not desc_col_idx:
            desc_col_idx = 1
            header_row_idx = 1

        # Шаг Б: Поиск шапки матрицы недель (W1, W2, W3... или 1, 2, 3...)
        # Карта: col_index -> week_number
        week_col_map = {}
        for r in range(max(1, header_row_idx - 3), min(header_row_idx + 3, ws.max_row + 1)):
            for c in range(desc_col_idx + 1, ws.max_column + 1):
                val = str(ws.cell(row=r, column=c).value or "").strip().upper()
                # Распознаем формат 'W1', 'W01', 'WEEK 1' или просто числовые значения недель
                if val.startswith("W") and val[1:].isdigit():
                    week_col_map[c] = int(val[1:])
                elif val.isdigit() and int(val) < 200:
                    week_col_map[c] = int(val)
            if len(week_col_map) >= 3:
                break

        # Если не нашли явные 'W1', ищем любые закрашенные колонки матрицы
        if not week_col_map:
            # Находим первую активную матричную колонку
            start_matrix_col = None
            for c in range(desc_col_idx + 5, ws.max_column + 1):
                for r_check in range(header_row_idx + 1, min(header_row_idx + 50, ws.max_row + 1)):
                    cell = ws.cell(row=r_check, column=c)
                    if cell.value or (cell.fill and cell.fill.fill_type and cell.fill.fill_type != 'none'):
                        start_matrix_col = c
                        break
                if start_matrix_col:
                    break
            
            if start_matrix_col:
                for c in range(start_matrix_col, ws.max_column + 1):
                    week_col_map[c] = (c - start_matrix_col) + 1

        if not week_col_map:
            continue

        # Шаг В: Считывание задач и закрашенных недель
        for r in range(header_row_idx + 1, ws.max_row + 1):
            desc_val = ws.cell(row=r, column=desc_col_idx).value
            if not desc_val or str(desc_val).strip() == "":
                continue
                
            task_desc = str(desc_val).strip()
            
            min_w = None
            max_w = None
            
            for c, w_num in week_col_map.items():
                cell = ws.cell(row=r, column=c)
                fill = cell.fill
                
                has_fill = False
                if fill and fill.fill_type and fill.fill_type != 'none':
                    fg = getattr(fill.start_color, 'rgb', None) or getattr(fill.start_color, 'theme', None)
                    if fg and str(fg) not in ['00000000', 'FFFFFFFF', '0']:
                        has_fill = True
                
                has_val = cell.value is not None and str(cell.value).strip() != ""
                
                if has_fill or has_val:
                    if min_w is None or w_num < min_w:
                        min_w = w_num
                    if max_w is None or w_num > max_w:
                        max_w = w_num
            
            # Если нет отметок в матрице, по умолчанию ставим неделе W1
            if min_w is None:
                min_w = 1
                max_w = 1
                
            duration_weeks = (max_w - min_w) + 1
            
            # Расчет точных дат:
            # Неделя 1 (W1) начинается точно в project_start_date
            task_start_date = project_start_date + timedelta(days=(min_w - 1) * 7)
            task_end_date = project_start_date + timedelta(days=max_w * 7 - 1)
            
            tasks.append({
                "phase": phase_label,
                "description": task_desc,
                "start_week_num": min_w,
                "end_week_num": max_w,
                "duration_weeks": duration_weeks,
                "start_date": task_start_date,
                "end_date": task_end_date
            })

    if not tasks:
        return None, None, None

    processed_tasks = []
    for idx, t in enumerate(tasks, 1):
        t["id"] = idx
        processed_tasks.append(t)

    df_tasks = pd.DataFrame(processed_tasks)
    
    # Сетка всех недель от W1 до максимальной недели проекта
    min_project_w = df_tasks["start_week_num"].min()
    max_project_w = df_tasks["end_week_num"].max()
    
    weeks_info = []
    for w_num in range(min_project_w, max_project_w + 1):
        w_start = project_start_date + timedelta(days=(w_num - 1) * 7)
        w_end = w_start + timedelta(days=6)
        has_holiday = check_week_has_holidays(w_start, w_end, ru_holidays)
        
        weeks_info.append({
            "week_num": w_num,
            "week_label": f"W{w_num}",
            "start_date": w_start,
            "end_date": w_end,
            "has_holiday": has_holiday
        })

    return df_tasks, weeks_info, milestones

# -----------------------------------------------------------------------------
# 4. ФОРМИРОВАНИЕ ИТОГОВОГО EXCEL
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

    headers = ["№", "Фаза (Phase)", "Описание действия (Description)", "Длит. (нед)", "Старт W", "Финиш W", "Дата начала", "Дата окончания"]
    for col_idx, h in enumerate(headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.font = font_header
        cell.fill = fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        cell.border = thin_border

    start_week_col = len(headers) + 1
    for i, w in enumerate(weeks_info):
        col_idx = start_week_col + i
        label = f"{w['week_label']} 🎉" if w["has_holiday"] else w["week_label"]
        cell = ws.cell(row=1, column=col_idx, value=label)
        cell.font = Font(name="Arial", size=9, bold=True, color="9C0006" if w["has_holiday"] else "FFFFFF")
        cell.fill = fill_holiday if w["has_holiday"] else fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
        ws.column_dimensions[get_column_letter(col_idx)].width = 4

    for row_idx, task in df_tasks.iterrows():
        r = row_idx + 2
        vals = [
            task["id"],
            task["phase"],
            task["description"],
            task["duration_weeks"],
            f"W{task['start_week_num']}",
            f"W{task['end_week_num']}",
            task["start_date"].strftime("%d.%m.%Y"),
            task["end_date"].strftime("%d.%m.%Y")
        ]
        
        for c_idx, val in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c_idx, value=val)
            cell.font = font_body
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center" if c_idx in [1, 4, 5, 6, 7, 8] else "left", vertical="center")

        # Отрисовка баров Ганта
        for w_idx, w in enumerate(weeks_info):
            c_idx = start_week_col + w_idx
            cell = ws.cell(row=r, column=c_idx)
            cell.border = thin_border
            
            current_w_num = w["week_num"]
            if task['start_week_num'] <= current_w_num <= task['end_week_num']:
                cell.fill = fill_gantt_bar
                cell.value = "█"
                cell.font = Font(name="Arial", size=9, color="2F5597")
                cell.alignment = Alignment(horizontal="center", vertical="center")

    base_widths = [6, 15, 50, 11, 10, 10, 13, 13]
    for idx, w in enumerate(base_widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    ws.freeze_panes = "I2"

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
# 5. ИНТЕРФЕЙС STREAMLIT
# -----------------------------------------------------------------------------
st.sidebar.header("⚙️ Входные параметры")

uploaded_file = st.sidebar.file_uploader(
    "Загрузите PLANT_MASTER_SCHEDULE P25077.xlsx", 
    type=["xlsx"]
)

start_date_input = st.sidebar.date_input(
    "Дата начала проекта (Неделя W1)",
    value=datetime.date.today(),
    format="DD.MM.YYYY"
)

if uploaded_file is not None:
    file_bytes = uploaded_file.getvalue()
    
    with st.spinner("Агент сканирует заголовки недель и рассчитывает даты..."):
        df_tasks, weeks_info, milestones = process_excel_schedule(file_bytes, start_date_input)
    
    if df_tasks is not None and not df_tasks.empty:
        c1, c2, c3, c4 = st.columns(4)
        c1.metric("Всего задач", len(df_tasks))
        c2.metric("Общий горизонт (нед)", len(weeks_info))
        c3.metric("Старт проекта", df_tasks["start_date"].min().strftime("%d.%m.%Y"))
        c4.metric("Окончание проекта", df_tasks["end_date"].max().strftime("%d.%m.%Y"))

        st.markdown("---")

        st.subheader("📊 Интерактивный SMS-график (Гант)")
        fig = px.timeline(
            df_tasks,
            x_start="start_date",
            x_end="end_date",
            y="description",
            color="phase",
            hover_data=["duration_weeks", "start_week_num", "end_week_num"],
            labels={"description": "Описание действия", "phase": "Фаза", "start_week_num": "Старт W", "end_week_num": "Финиш W"}
        )
        fig.update_yaxes(autorange="reversed")
        fig.update_layout(height=min(800, 150 + len(df_tasks) * 25))
        st.plotly_chart(fig, use_container_width=True)

        with st.expander("📋 Детализация календарного плана (Таблица)"):
            st.dataframe(
                df_tasks[["id", "phase", "description", "start_week_num", "end_week_num", "duration_weeks", "start_date", "end_date"]],
                use_container_width=True
            )

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
        st.error("Не удалось прочитать задачи или шапку недель в файле.")
else:
    st.info("👈 Загрузите исходный Excel-файл на панели слева.")
