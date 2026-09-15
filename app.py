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
на основе шаблона **PLANT_MASTER_SCHEDULE P25077** и **Производственного календаря РФ**.
""")

# ---------------------------------------------------------
# Строгие константы ТЗ
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

START_MILESTONE_SHEET = "START PROJECT TOGF-ENG-007-02"

# ---------------------------------------------------------
# Вспомогательные функции для работы с цветом и датами
# ---------------------------------------------------------
def is_cell_colored(cell):
    """
    Проверяет, имеет ли ячейка заливку цветом (Condition 5: определение Duration weeks).
    """
    if not cell.fill or cell.fill.fill_type in (None, "none"):
        return False
    
    color = cell.fill.start_color
    if color:
        if color.rgb:
            rgb_val = str(color.rgb).upper()
            # Игнорируем стандартную белую или прозрачную заливку
            if rgb_val in ["00000000", "FFFFFFFF", "00FFFFFF", "FFFFFF", "000000"]:
                return False
            return True
        if color.theme is not None:
            return True
        if color.indexed is not None and color.indexed != 64:
            return True
            
    return False

@st.cache_data
def get_russian_holidays(start_year=2024, end_year=2030):
    """
    Формирует список государственных праздников РФ (библиотека holidays RU / КонсультантПлюс).
    """
    ru_holidays = set()
    for yr in range(start_year, end_year + 1):
        for h_date in holidays.RU(years=yr).keys():
            ru_holidays.add(h_date)
    return ru_holidays

def add_working_days(start_date, num_days, ru_holidays):
    """
    Прибавляет рабочие дни с учётом обычных выходных (СБ, ВС) и праздников РФ.
    """
    current_date = start_date
    added = 0
    while added < num_days:
        current_date += timedelta(days=1)
        if current_date.weekday() < 5 and current_date not in ru_holidays:
            added += 1
    return current_date

# ---------------------------------------------------------
# Парсинг и обработка шаблона Excel
# ---------------------------------------------------------
def process_excel_schedule(uploaded_file):
    wb = openpyxl.load_workbook(uploaded_file, data_only=True)
    ru_holidays = get_russian_holidays(2024, 2030)
    
    # 1. Извлечение стартовой даты вехи (Условие 1)
    start_sheet_name = None
    for name in wb.sheetnames:
        if "START PROJECT" in name.upper():
            start_sheet_name = name
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
        st.warning(f"Дата вехи не найдена на листе {START_MILESTONE_SHEET}. Установлена текущая дата: {project_start_date}")

    # 2. Обработка фазовых вкладок в строгой последовательности (Условие 2)
    tasks = []
    current_date = project_start_date

    for target_phase in PHASE_SHEETS:
        matched_sheet = None
        for name in wb.sheetnames:
            if name.strip().replace(" ", "").upper() in target_phase.strip().replace(" ", "").upper():
                matched_sheet = name
                break
                
        if not matched_sheet:
            continue
            
        ws = wb[matched_sheet]
        
        # Нахождение колонки DESCRIPTION
        desc_col_idx = None
        header_row_idx = 1
        
        for r_idx in range(1, min(15, ws.max_row + 1)):
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

        # Обход строк задач
        for r_idx in range(header_row_idx + 1, ws.max_row + 1):
            desc_val = ws.cell(row=r_idx, column=desc_col_idx).value
            if not desc_val or str(desc_val).strip() == "":
                continue
                
            desc_str = str(desc_val).strip()
            
            # Условие 3: Игнорирование системных колонок/заголовков
            if any(ign.lower() in desc_str.lower() for ign in IGNORED_COLUMNS):
                continue

            # Условие 5: Подсчет Duration weeks по цветной заливке ячеек в этой строке
            colored_weeks_count = 0
            for c_idx in range(desc_col_idx + 1, ws.max_column + 1):
                cell = ws.cell(row=r_idx, column=c_idx)
                if is_cell_colored(cell):
                    colored_weeks_count += 1

            # Если заливки нет, по умолчанию 1 неделя
            duration_weeks = colored_weeks_count if colored_weeks_count > 0 else 1
            duration_working_days = duration_weeks * 5
            
            # Корректировка даты начала (не должна выпадать на выходной/праздник)
            while current_date.weekday() >= 5 or current_date in ru_holidays:
                current_date += timedelta(days=1)

            task_start = current_date
            task_end = add_working_days(task_start, duration_working_days, ru_holidays)
            
            tasks.append({
                "Phase": matched_sheet,
                "Description": desc_str,
                "Duration_Weeks": duration_weeks,
                "Duration_Days": duration_working_days,
                "Start_Date": task_start,
                "End_Date": task_end,
            })
            
            # Переход к следующей задаче
            current_date = task_end + timedelta(days=1)
            while current_date.weekday() >= 5 or current_date in ru_holidays:
                current_date += timedelta(days=1)

    return pd.DataFrame(tasks), project_start_date

# ---------------------------------------------------------
# Генератор файла Excel с диаграммой Ганта на листе
# ---------------------------------------------------------
def create_excel_with_embedded_gantt(df_tasks, project_start_date):
    output = io.BytesIO()
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SMS Master Schedule"
    ws.views.sheetView[0].showGridLines = True

    # Стилизация
    HEADER_FILL = PatternFill(start_color="1F4E79", end_color="1F4E79", fill_type="solid")
    WEEK_FILL = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    GANTT_FILL = PatternFill(start_color="5B9BD5", end_color="5B9BD5", fill_type="solid")
    ALT_ROW_FILL = PatternFill(start_color="F2F4F7", end_color="F2F4F7", fill_type="solid")
    
    HEADER_FONT = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    REG_FONT = Font(name="Calibri", size=10)
    
    BORDER_THIN = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # Заголовок документа
    ws.cell(row=1, column=1, value="ГРАФИК ЗАПУСКА В СЕРИЙНОЕ ПРОИЗВОДСТВО (SMS)").font = Font(name="Calibri", size=14, bold=True, color="1F4E79")
    ws.cell(row=2, column=1, value=f"Старт проекта: {project_start_date.strftime('%d.%m.%Y')} | Календарь РФ").font = Font(name="Calibri", size=10, italic=True)

    base_headers = ["№", "Фаза проекта", "Описание действия (DESCRIPTION)", "Длительность (нед.)", "Дата начала", "Дата окончания"]
    start_row = 4

    for col_idx, text in enumerate(base_headers, start=1):
        c = ws.cell(row=start_row, column=col_idx, value=text)
        c.fill = HEADER_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # Таймлайн по неделям
    total_weeks = df_tasks["Duration_Weeks"].sum()
    gantt_start_col = len(base_headers) + 1

    for w_idx in range(total_weeks):
        col = gantt_start_col + w_idx
        c = ws.cell(row=start_row, column=col, value=f"W{w_idx + 1}")
        c.fill = WEEK_FILL
        c.font = HEADER_FONT
        c.alignment = Alignment(horizontal="center", vertical="center")
        ws.column_dimensions[get_column_letter(col)].width = 5

    # Заполнение табличной части и закрашивание ячеек Ганта
    current_w_offset = 0
    for idx, row in df_tasks.iterrows():
        r = start_row + 1 + idx
        
        ws.cell(row=r, column=1, value=idx + 1).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=2, value=row["Phase"])
        ws.cell(row=r, column=3, value=row["Description"])
        ws.cell(row=r, column=4, value=row["Duration_Weeks"]).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=5, value=row["Start_Date"].strftime("%d.%m.%Y")).alignment = Alignment(horizontal="center")
        ws.cell(row=r, column=6, value=row["End_Date"].strftime("%d.%m.%Y")).alignment = Alignment(horizontal="center")

        for col_idx in range(1, len(base_headers) + 1):
            c = ws.cell(row=r, column=col_idx)
            c.font = REG_FONT
            c.border = BORDER_THIN
            if idx % 2 == 1:
                c.fill = ALT_ROW_FILL

        # Подсветка полосы Ганта в соответствии с длительностью в неделях
        dur_w = row["Duration_Weeks"]
        for w in range(dur_w):
            gc = ws.cell(row=r, column=gantt_start_col + current_w_offset + w)
            gc.fill = GANTT_FILL
            gc.border = BORDER_THIN
            
        current_w_offset += dur_w

    # Ширина столбцов
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 28
    ws.column_dimensions['C'].width = 52
    ws.column_dimensions['D'].width = 16
    ws.column_dimensions['E'].width = 14
    ws.column_dimensions['F'].width = 14

    wb.save(output)
    output.seek(0)
    return output

# ---------------------------------------------------------
# Streamlit UI
# ---------------------------------------------------------
uploaded_file = st.file_uploader(
    "Загрузите исходный файл Excel (PLANT_MASTER_SCHEDULE P25077.xlsx)",
    type=["xlsx"]
)

if uploaded_file is not None:
    with st.spinner("Анализ цветной заливки ячеек, извлечение вех и расчёт дат..."):
        df_tasks, project_start_date = process_excel_schedule(uploaded_file)

    if not df_tasks.empty:
        st.success(f"SMS-график успешно построен! Базовая дата из вехи: **{project_start_date.strftime('%d.%m.%Y')}**")

        # Дашборд
        m1, m2, m3, m4 = st.columns(4)
        m1.metric("Всего этапов", len(df_tasks))
        m2.metric("Сумма недель", df_tasks["Duration_Weeks"].sum())
        m3.metric("Дата старта", df_tasks["Start_Date"].min().strftime("%d.%m.%Y"))
        m4.metric("Серийный запуск", df_tasks["End_Date"].max().strftime("%d.%m.%Y"))

        st.markdown("---")
        st.subheader("📊 Интерактивная Диаграмма Ганта (Plotly)")

        fig = px.timeline(
            df_tasks,
            x_start="Start_Date",
            x_end="End_Date",
            y="Description",
            color="Phase",
            title="График запуска проекта в серийное производство (SMS)",
            hover_data=["Duration_Weeks", "Start_Date", "End_Date"]
        )
        fig.update_yaxes(autorange="reversed")
        fig.update_layout(
            height=320 + len(df_tasks) * 25,
            xaxis_title="Временная шкала (с учётом выходных и праздников РФ)",
            yaxis_title="Действие (DESCRIPTION)",
            legend_title="Фаза проекта"
        )
        st.plotly_chart(fig, use_container_width=True)

        st.subheader("📋 Сформированный календарный план")
        display_df = df_tasks.copy()
        display_df["Start_Date"] = display_df["Start_Date"].apply(lambda x: x.strftime("%d.%m.%Y"))
        display_df["End_Date"] = display_df["End_Date"].apply(lambda x: x.strftime("%d.%m.%Y"))
        
        st.dataframe(
            display_df[["Phase", "Description", "Duration_Weeks", "Start_Date", "End_Date"]],
            use_container_width=True
        )

        excel_data = create_excel_with_embedded_gantt(df_tasks, project_start_date)
        st.download_button(
            label="📥 Скачать Excel с Диаграммой Ганта на листе",
            data=excel_data,
            file_name=f"SMS_Schedule_P25077_{datetime.date.today().strftime('%Y%m%d')}.xlsx",
            mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
        )
    else:
        st.warning("Не удалось найти действия на указанных фазовых листах.")
else:
    st.info("Пожалуйста, загрузите файл шаблона `PLANT_MASTER_SCHEDULE P25077.xlsx`.")
