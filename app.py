import io
import datetime
from datetime import timedelta
import pandas as pd
import numpy as np
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter
import plotly.express as px
import streamlit as st
import holidays

# -----------------------------------------------------------------------------
# 1. КОНФИГУРАЦИЯ СТРАНИЦЫ И СТИЛИ
# -----------------------------------------------------------------------------
st.set_page_config(
    page_title="Генератор SMS-графика проекта (TOGF-ENG)",
    page_icon="📅",
    layout="wide"
)

st.title("📅 Генератор SMS-графика проекта запуска (TOGF-ENG)")
st.caption("Автоматическое формирование диаграммы Ганта из шаблона Excel с учетом производственного календаря РФ")

# -----------------------------------------------------------------------------
# 2. ПОМОЩНИКИ И КАЛЕНДАРЬ РФ
# -----------------------------------------------------------------------------
@st.cache_data
def get_ru_holidays(year_start: int, year_end: int):
    """Получение праздничных дней РФ за указанные года."""
    ru_holidays = set()
    for yr in range(year_start, year_end + 1):
        for d, name in holidays.RU(years=yr).items():
            ru_holidays.add(d)
    return ru_holidays

def check_week_has_holidays(week_start: datetime.date, week_end: datetime.date, ru_holidays: set) -> bool:
    """Проверяет, выпадает ли государственный праздник РФ на указанную неделю."""
    curr = week_start
    while curr <= week_end:
        if curr in ru_holidays:
            return True
        curr += timedelta(days=1)
    return False

# -----------------------------------------------------------------------------
# 3. ОСНОВНОЙ АЛГОРИТМ ПАРСИНГА И РАСЧЕТА
# -----------------------------------------------------------------------------
def process_excel_schedule(file_bytes, project_start_date: datetime.date):
    wb = openpyxl.load_workbook(filename=io.BytesIO(file_bytes), data_only=True)
    
    # --- Шаг 1: Извлечение вех проекта (Milestones) ---
    milestones = []
    milestone_sheet_name = "START PROJECT TOGF-ENG-007-02"
    if milestone_sheet_name in wb.sheetnames:
        ws_m = wb[milestone_sheet_name]
        for row in ws_m.iter_rows(values_only=True):
            vals = [str(cell).strip() for cell in row if cell is not None and str(cell).strip() != ""]
            if vals:
                milestones.append(" | ".join(vals[:3]))
    
    # --- Шаг 2 & 3: Парсинг задач и расчет длительностей ---
    target_sheets = [
        ("PMSPR TOGF-ENG-008-06 Phase2", "Phase 2"),
        ("PMSPR TOGF-ENG-008-06 Phase 3", "Phase 3"),
        ("PMSPR TOGF-ENG-008-06 Phase4;5", "Phase 4;5")
    ]
    
    tasks = []
    
    for sheet_name, phase_label in target_sheets:
        if sheet_name not in wb.sheetnames:
            continue
            
        ws = wb[sheet_name]
        
        # Поиск колонки DESCRIPTION
        desc_col_idx = None
        header_row = 1
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
            desc_col_idx = 1 # Запасной вариант, если заголовок не найден
            
        # Анализ строк задач
        # Колонка AM — это 39-я колонка в Excel (A=1, ..., AM=39)
        start_matrix_col = 39 
        
        for r in range(header_row + 1, ws.max_row + 1):
            desc_val = ws.cell(row=r, column=desc_col_idx).value
            if not desc_val or str(desc_val).strip() == "":
                continue
                
            task_desc = str(desc_val).strip()
            
            # Подсчет длительности (сканирование ячеек начиная с AM)
            duration_weeks = 0
            max_check_col = max(start_matrix_col + 100, ws.max_column + 1)
            
            for c in range(start_matrix_col, max_check_col):
                cell = ws.cell(row=r, column=c)
                fill = cell.fill
                
                has_fill = False
                # Проверка заливки цветным фоном
                if fill and fill.fill_type and fill.fill_type != 'none':
                    fg_color = getattr(fill.start_color, 'rgb', None) or getattr(fill.start_color, 'theme', None)
                    if fg_color and str(fg_color) not in ['00000000', 'FFFFFFFF', '0']:
                        has_fill = True
                
                # Проверка наличия значения в ячейке
                has_val = cell.value is not None and str(cell.value).strip() != ""
                
                if has_fill or has_val:
                    duration_weeks += 1
                else:
                    # Если до этого были залитые ячейки, а сейчас пустая — прерываем последовательность
                    if duration_weeks > 0:
                        break
            
            # Правило минимальной длительности
            if duration_weeks == 0:
                duration_weeks = 1
                
            tasks.append({
                "phase": phase_label,
                "description": task_desc,
                "duration_weeks": duration_weeks
            })

    if not tasks:
        st.error("Не удалось извлечь задачи из указанных вкладок Excel. Проверьте структуру файла.")
        return None, None, None

    # --- Шаг 4: Расчет календарных дат ---
    current_start = project_start_date
    processed_tasks = []
    
    # Кэш праздников
    ru_holidays = get_ru_holidays(project_start_date.year, project_start_date.year + 3)
    
    for idx, t in enumerate(tasks, 1):
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
        # Следующая задача начинается на следующий день после окончания предыдущей
        current_start = task_end + timedelta(days=1)

    df_tasks = pd.DataFrame(processed_tasks)
    
    # Определение структуры недель проекта
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
# 4. ФУНКЦИЯ СБОРКИ EXCEL-ОТЧЕТА С ГАНТОМ ПО ЯЧЕЙКАМ
# -----------------------------------------------------------------------------
def generate_excel_report(df_tasks, weeks_info, milestones):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SMS Schedule"
    
    # Включение отображения сетки
    ws.views.sheetView[0].showGridLines = True
    
    # Стили
    font_header = Font(name="Arial", size=9, bold=True, color="FFFFFF")
    font_body = Font(name="Arial", size=9)
    font_bold = Font(name="Arial", size=9, bold=True)
    
    fill_header = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    fill_phase = PatternFill(start_color="D9E1F2", end_color="D9E1F2", fill_type="solid")
    fill_gantt_bar = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    fill_holiday_week = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid") # Светло-красный
    
    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # 1. Заголовок таблицы
    headers = ["№", "Фаза (Phase)", "Описание действия (Description)", "Длит. (нед)", "Дата начала", "Дата окончания"]
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
        cell.fill = fill_holiday_week if w["has_holiday"] else fill_header
        cell.alignment = Alignment(horizontal="center", vertical="center")
        cell.border = thin_border
        
        # ТЗ Требование 4: Ширина столбцов недель ровно 4
        col_letter = get_column_letter(col_idx)
        ws.column_dimensions[col_letter].width = 4

    # 2. Данные задач
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
            if c_idx in [1, 4, 5, 6]:
                cell.alignment = Alignment(horizontal="center", vertical="center")
            else:
                cell.alignment = Alignment(horizontal="left", vertical="center")

        # Отрисовка бара Ганта
        for w_idx, w in enumerate(weeks_info):
            c_idx = start_week_col + w_idx
            cell = ws.cell(row=r, column=c_idx)
            cell.border = thin_border
            
            # Пересечение сроков задачи с неделей
            if not (task["end_date"] < w["start_date"] or task["start_date"] > w["end_date"]):
                cell.fill = fill_gantt_bar
                cell.value = "█"
                cell.font = Font(name="Arial", size=9, color="2F5597")
                cell.alignment = Alignment(horizontal="center", vertical="center")

    # Автоширина базовых колонок
    base_widths = [6, 15, 50, 12, 14, 14]
    for idx, w in enumerate(base_widths, 1):
        ws.column_dimensions[get_column_letter(idx)].width = w

    # Закрепление шапки и колонок
    ws.freeze_panes = "G2"

    # 3. Лист с вехами проекта (Справочно)
    if milestones:
        ws_m = wb.create_sheet(title="Milestones")
        ws_m.views.sheetView[0].showGridLines = True
        ws_m.cell(row=1, column=1, value="Ключевые вехи проекта (START PROJECT TOGF-ENG-007-02)").font = font_bold
        for idx, m_text in enumerate(milestones, 2):
            ws_m.cell(row=idx, column=1, value=m_text).font = font_body

    output = io.BytesIO()
    wb.save(output)
    output.seek(0)
    return output

# -----------------------------------------------------------------------------
# 5. ПОЛЬЗОВАТЕЛЬСКИЙ ИНТЕРФЕЙС STREAMLIT
# -----------------------------------------------------------------------------
st.sidebar.header("⚙️ Параметры запуска")

uploaded_file = st.sidebar.file_content = st.sidebar.file_uploader(
    "Загрузите шаблон Excel", 
    type=["xlsx"],
    help="Файл PLANT_MASTER_SCHEDULE P25077.xlsx"
)

# Запрос даты старта (ТЗ Раздел 2)
start_date_input = st.sidebar.date_input(
    "Дата начала проекта",
    value=datetime.date.today(),
    format="DD.MM.YYYY"
)

if uploaded_file is not None:
    if st.sidebar.button("🚀 Сформировать график", type="primary"):
        with st.spinner("ИИ-агент обрабатывает файл и рассчитывает календарные сроки..."):
            file_bytes = uploaded_file.read()
            df_tasks, weeks_info, milestones = process_excel_schedule(file_bytes, start_date_input)
            
            if df_tasks is not None:
                st.session_state["df_tasks"] = df_tasks
                st.session_state["weeks_info"] = weeks_info
                st.session_state["milestones"] = milestones
                st.success("Расчет успешно завершен!")

if "df_tasks" in st.session_state:
    df_tasks = st.session_state["df_tasks"]
    weeks_info = st.session_state["weeks_info"]
    milestones = st.session_state["milestones"]

    # --- МЕТРИКИ ---
    col1, col2, col3, col4 = st.columns(4)
    col1.metric("Всего задач", len(df_tasks))
    col2.metric("Длительность (нед)", df_tasks["duration_weeks"].sum())
    col3.metric("Дата начала", df_tasks["start_date"].min().strftime("%d.%m.%Y"))
    col4.metric("Дата окончания", df_tasks["end_date"].max().strftime("%d.%m.%Y"))

    st.markdown("---")

    # --- ВИЗУАЛИЗАЦИЯ PLOTLY GANTT ---
    st.subheader("📊 Интерактивная Диаграмма Ганта (Plotly)")
    
    fig = px.timeline(
        df_tasks,
        x_start="start_date",
        x_end="end_date",
        y="description",
        color="phase",
        hover_data=["duration_weeks", "phase"],
        labels={"description": "Описание действия", "phase": "Фаза", "duration_weeks": "Длительность (нед)"},
        title="Календарный план-график проекта"
    )
    
    fig.update_yaxes(autorange="reversed") # Порядок задач сверху вниз
    fig.update_layout(
        height=min(800, 100 + len(df_tasks) * 25),
        xaxis_title="Дата / Недели",
        yaxis_title="",
        legend_title="Фаза проекта"
    )
    
    st.plotly_chart(fig, use_container_width=True)

    # --- ТАБЛИЦА С ДАННЫМИ ---
    with st.expander("📋 Просмотр таблицы рассчитанных задач", expanded=False):
        st.dataframe(
            df_tasks[["id", "phase", "description", "duration_weeks", "start_date", "end_date"]],
            use_container_width=True
        )

    # --- СКАЧИВАНИЕ ФАЙЛА EXCEL ---
    st.markdown("---")
    st.subheader("📥 Экспорт результата")
    
    excel_data = generate_excel_report(df_tasks, weeks_info, milestones)
    
    st.download_button(
        label="💾 Скачать итоговый Excel-файл с Диаграммой Ганта",
        data=excel_data,
        file_name=f"SMS_Schedule_{start_date_input.strftime('%Y%m%d')}.xlsx",
        mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
        type="primary"
    )
else:
    st.info("👈 Загрузите файл `PLANT_MASTER_SCHEDULE P25077.xlsx` на панели слева и нажмите **'Сформировать график'**.")
