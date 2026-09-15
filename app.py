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
    page_title="Генератор SMS-графика проекта (TOGF-ENG)",
    page_icon="📅",
    layout="wide"
)

st.title("📅 Генератор SMS-графика проекта запуска (TOGF-ENG)")
st.caption("Автоматическое формирование диаграммы Ганта из шаблона Excel с учетом производственного календаря РФ")

# -----------------------------------------------------------------------------
# 2. ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ И КАЛЕНДАРЬ
# -----------------------------------------------------------------------------
@st.cache_data
def get_ru_holidays_set(year_start: int, year_end: int):
    """Получение множества дат (datetime.date) праздников РФ."""
    ru_holidays = set()
    for yr in range(year_start, year_end + 1):
        for d in holidays.RU(years=yr).keys():
            if isinstance(d, datetime.datetime):
                ru_holidays.add(d.date())
            else:
                ru_holidays.add(d)
    return ru_holidays

def check_week_has_holidays(week_start: datetime.date, week_end: datetime.date, ru_holidays: set) -> bool:
    """Проверка наличия государственного праздника на неделе."""
    curr = week_start
    while curr <= week_end:
        if curr in ru_holidays:
            return True
        curr += timedelta(days=1)
    return False

# -----------------------------------------------------------------------------
# 3. ПАРСИНГ И РАСЧЕТ ДАТ
# -----------------------------------------------------------------------------
def process_excel_schedule(file_bytes, project_start_date: datetime.date):
    wb = openpyxl.load_workbook(filename=io.BytesIO(file_bytes), data_only=True)
    
    # --- Извлечение вех ---
    milestones = []
    milestone_sheet_name = "START PROJECT TOGF-ENG-007-02"
    if milestone_sheet_name in wb.sheetnames:
        ws_m = wb[milestone_sheet_name]
        for row in ws_m.iter_rows(values_only=True):
            vals = [str(cell).strip() for cell in row if cell is not None and str(cell).strip() != ""]
            if vals:
                milestones.append(" | ".join(vals[:3]))
    
    # --- Парсинг задач ---
    target_sheets = [
        ("PMSPR TOGF-ENG-008-06 Phase2", "Phase 2"),
        ("PMSPR TOGF-ENG-008-06 Phase 3", "Phase 3"),
        ("PMSPR TOGF-ENG-008-06 Phase4;5", "Phase 4;5")
    ]
    
    tasks = []
    start_matrix_col = 39 # Колонка AM (39)
    
    for sheet_name, phase_label in target_sheets:
        if sheet_name not in wb.sheetnames:
            continue
            
        ws = wb[sheet_name]
        desc_col_idx = None
        header_row = 1
        
        # Поиск колонки DESCRIPTION
        for r in range(1, min(15, ws.max_row + 1)):
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
            max_check_col = max(start_matrix_col + 100, ws.max_column + 1)
            
            for c in range(start_matrix_col, max_check_col):
                cell = ws.cell(row=r, column=c)
                fill = cell.fill
                
                has_fill = False
                if fill and fill.fill_type and fill.fill_type != 'none':
                    fg_color = getattr(fill.start_color, 'rgb', None) or getattr(fill.start_color, 'theme', None)
                    if fg_color and str(fg_color) not in ['00000000', 'FFFFFFFF', '0']:
                        has_fill = True
                
                has_val = cell.value is not None and str(cell.value).strip() != ""
                
                if has_fill or has_val:
                    duration_weeks += 1
                else:
                    if duration_weeks > 0:
                        break
            
            if duration_weeks == 0:
                duration_weeks = 1
                
            tasks.append({
                "phase": phase_label,
                "description": task_desc,
                "duration_weeks": duration_weeks
            })

    if not tasks:
        return None, None, None

    # --- Точный расчет календарных дат ---
    ru_holidays = get_ru_holidays_set(project_start_date.year, project_start_date.year + 4)
    current_start = project_start_date
    processed_tasks = []
    
    for idx, t in enumerate(tasks, 1):
        # Длительность задачи в днях
        duration_days = t["duration_weeks"] * 7
        task_end = current_start + timedelta(days=duration_days - 1)
        
        processed_tasks.append({
            "id": idx,
            "phase": t["phase"],
            "description": t["description"],
            "duration_weeks": t["duration_weeks"],
            "start_date": current_start,
            "end_date": task_end,
        })
        current_start = task_end + timedelta(days=1)

    df_tasks = pd.DataFrame(processed_tasks)
    
    # --- Построение структуры недель ---
    max_end_date = df_tasks["end_date"].max()
    weeks_info = []
    w_start = project_start_date
    w_num = 1
    
    while w_start <= max_end_date + timedelta(days=7):
        w_end = w_start + timedelta(days=6)
        has_holiday = check_week_has_holidays(w_start, w_end, ru_holidays)
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
# 4. ГЕНЕРАЦИЯ EXCEL (С ДИАГРАММОЙ ГАНТА ПО ЯЧЕЙКАМ)
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
    fill_holiday_week = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    headers = ["№", "Фаза (Phase)", "Описание действия (Description)", "Длит. (нед)", "Дата начала", "Дата окончания"]
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
        cell.fill = fill_holiday_week if w["has_holiday"] else fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
        # Ширина столбца недели = 4
        ws.column_dimensions[get_column_letter(col_idx)].width = 4

    for row_idx, task in df_tasks.iterrows():
        r = row_idx + 2
        vals = [
            task["id"],
            task["phase"],
            task["description"],
            task["duration_weeks"],
            task["start_date"].strftime("%d.%m.%Y"),
            task["end_date"].strftime("%d.%m.%Y")
        ]
        
        for c_idx, val in enumerate(vals, 1):
            cell = ws.cell(row=r, column=c_idx, value=val)
            cell.font = font_body
            cell.border = thin_border
            cell.alignment = Alignment(horizontal="center" if c_idx in [1, 4, 5, 6] else "left", vertical="center")

        for w_idx, w in enumerate(weeks_info):
            c_idx = start_week_col + w_idx
            cell = ws.cell(row=r, column=c_idx)
            cell.border = thin_border
            
            if not (task["end_date"] < w["start_date"] or task["start_date"] > w["end_date"]):
                cell.fill = fill_gantt_bar
                cell.value = "█"
                cell.font = Font(name="Arial", size=9, color="2F5597")
                cell.alignment = Alignment(horizontal="center", vertical="center")

    base_widths = [6, 15, 50, 12, 14, 14]
    for idx, w in enumerate(base_widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    ws.freeze_panes = "G2"

    if milestones:
        ws_m = wb.create_sheet(title="Milestones")
        ws_m.views.sheetView[0].showGridLines = True
        ws_m.cell(row=1, column=1, value="Ключевые вехи проекта").font = font_bold
        for idx, m_text in enumerate(milestones, 2):
            ws_m.cell(row=idx, column=1, value=m_text).font = font_body

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# -----------------------------------------------------------------------------
# 5. ИНТЕРФЕЙС И КНОПКА СКАЧИВАНИЯ
# -----------------------------------------------------------------------------
st.sidebar.header("⚙️ Параметры")

uploaded_file = st.sidebar.file_uploader(
    "Загрузите Excel-файл", 
    type=["xlsx"]
)

start_date_input = st.sidebar.date_input(
    "Дата начала проекта",
    value=datetime.date.today(),
    format="DD.MM.YYYY"
)

if uploaded_file is not None:
    # Автоматическая обработка при загрузке или смене даты
    file_bytes = uploaded_file.getvalue()
    df_tasks, weeks_info, milestones = process_excel_schedule(file_bytes, start_date_input)
    
    if df_tasks is not None:
        # 1. Метрики
        col1, col2, col3, col4 = st.columns(4)
        col1.metric("Задач", len(df_tasks))
        col2.metric("Длительность (нед)", df_tasks["duration_weeks"].sum())
        col3.metric("Старт", df_tasks["start_date"].min().strftime("%d.%m.%Y"))
        col4.metric("Финиш", df_tasks["end_date"].max().strftime("%d.%m.%Y"))

        st.markdown("---")

        # 2. График Plotly
        st.subheader("📊 Интерактивная Диаграмма Ганта")
        fig = px.timeline(
            df_tasks,
            x_start="start_date",
            x_end="end_date",
            y="description",
            color="phase",
            hover_data=["duration_weeks"],
            labels={"description": "Задача", "phase": "Фаза", "duration_weeks": "Недель"}
        )
        fig.update_yaxes(autorange="reversed")
        fig.update_layout(height=min(700, 150 + len(df_tasks) * 25))
        st.plotly_chart(fig, use_container_width=True)

        # 3. Кнопка скачивания Excel (Гарантированное отображение)
        st.markdown("---")
        st.subheader("📥 Скачать результаты")
        
        excel_file_data = generate_excel_report(df_tasks, weeks_info, milestones)
        
        st.download_button(
            label="💾 Скачать Excel-файл с диаграммой Ганта",
            data=excel_file_data,
            file_name=f"SMS_Schedule_{start_date_input.strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            type="primary",
            use_container_width=True
        )
    else:
        st.error("Ошибка парсинга файла. Проверьте наличие вкладок Phase2, Phase 3 или Phase4;5.")
else:
    st.info("👈 Загрузите файл `PLANT_MASTER_SCHEDULE P25077.xlsx` на панели слева.")
