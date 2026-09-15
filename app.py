import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from datetime import datetime, timedelta
import holidays
import io

# --- Настройка страницы ---
st.set_page_config(page_title="SMS Gantt Generator", page_icon="📊", layout="wide")
st.title("📊 Генератор SMS графика проекта запуска (TOGF-ENG)")
st.markdown("""
Инструмент формирует календарный план-график (Диаграмму Ганта) на основе шаблона `PLANT_MASTER_SCHEDULE P25077`.
- Учитывает вехи и описания из заданных вкладок.
- Игнорирует автоматические поля дат.
- Рассчитывает длительность (Duration weeks) по количеству залитых цветом ячеек (начиная с колонки AM).
- Учитывает государственные праздники РФ (помечает их в шапке графика).
""")

# --- Вспомогательные функции ---
def get_ru_holidays(years_range):
    """Возвращает множество дат государственных праздников РФ"""
    return set(holidays.RU(years=years_range))

def is_holiday_week(start_date, ru_holidays_set):
    """Проверяет, попадает ли государственный праздник (не выходной) на неделю, начинающуюся с start_date"""
    # Проверяем 7 дней вперед от начала недели
    for i in range(7):
        current_date = start_date + timedelta(days=i)
        if current_date in ru_holidays_set:
            return True
    return False

def count_colored_cells_in_row(ws, row_idx, start_col_str="AM", end_col_str="AZ"):
    """
    Считает количество ячеек в строке с заливкой (не белой и не пустой) 
    в заданном диапазоне колонок, что соответствует логике Duration weeks.
    """
    start_col = column_index_from_string(start_col_str)
    end_col = column_index_from_string(end_col_str)
    colored_count = 0
    
    for col in range(start_col, end_col + 1):
        cell = ws.cell(row=row_idx, column=col)
        # Проверяем наличие заливки, отличной от белой/прозрачной
        if cell.fill and cell.fill.fgColor:
            color_rgb = cell.fill.fgColor.rgb if isinstance(cell.fill.fgColor.rgb, str) else str(cell.fill.fgColor.rgb)
            # Игнорируем прозрачный ('00000000') и белый ('FFFFFF' или 'FFFFFFFF') цвета
            if color_rgb not in ['00000000', 'FFFFFF', 'FFFFFFFF']:
                # Дополнительно проверяем, что ячейка не совсем пустая (на случай артефактов форматирования)
                if cell.value is not None and str(cell.value).strip() != "":
                    colored_count += 1
    return colored_count

def find_column_index(ws, target_names, max_search=50):
    """Ищет индекс колонки по списку возможных названий"""
    for row in ws.iter_rows(min_row=1, max_row=5, max_col=max_search):
        for cell in row:
            if cell.value and isinstance(cell.value, str):
                if any(target.lower() in cell.value.lower() for target in target_names):
                    return cell.column
    return None

# --- Основной интерфейс ---
uploaded_file = st.file_uploader("Загрузите файл шаблона 'PLANT_MASTER_SCHEDULE P25077' (.xlsx)", type=["xlsx"])
project_start_date = st.date_input("Дата начала отсчета графика:", datetime.now())

if st.button("🚀 Сформировать SMS График", type="primary", disabled=not uploaded_file):
    if uploaded_file is not None:
        with st.spinner("Анализ файла и построение графика..."):
            try:
                # Загружаем книгу через openpyxl для сохранения форматирования и чтения цветов
                wb = openpyxl.load_workbook(uploaded_file, data_only=True)
                
                # 1. Сбор вех из "START PROJECT TOGF-ENG-007-02"
                milestones = []
                if "START PROJECT TOGF-ENG-007-02" in wb.sheetnames:
                    ws_milestones = wb["START PROJECT TOGF-ENG-007-02"]
                    # Ищем колонки с названиями вех (предполагаемые имена)
                    desc_col = find_column_index(ws_milestones, ["milestone", "веха", "name", "наименование", "description"])
                    if not desc_col:
                        desc_col = 1 # Fallback на первую колонку
                    
                    for row in range(2, ws_milestones.max_row + 1):
                        val = ws_milestones.cell(row=row, column=desc_col).value
                        if val and str(val).strip():
                            milestones.append(str(val).strip())
                
                # 2. Сбор описаний из Phase 2, 3, 4;5
                phase_tabs = [
                    "PMSPR TOGF-ENG-008-06 Phase2",
                    "PMSPR TOGF-ENG-008-06 Phase 3",
                    "PMSPR TOGF-ENG-008-06 Phase4;5"
                ]
                
                tasks_data = []
                
                for tab_name in phase_tabs:
                    if tab_name in wb.sheetnames:
                        ws = wb[tab_name]
                        # Ищем колонку DESCRIPTION
                        desc_col_idx = find_column_index(ws, ["description", "описание", "действие"])
                        
                        if desc_col_idx:
                            for row in range(2, ws.max_row + 1):
                                desc_val = ws.cell(row=row, column=desc_col_idx).value
                                if desc_val and str(desc_val).strip():
                                    # Считаем длительность по цветным ячейкам начиная с AM
                                    duration_weeks = count_colored_cells_in_row(ws, row, start_col_str="AM", end_col_str="AZ")
                                    
                                    # Если цвет не найден, но есть текст, ставим минимальную длительность 1 или 0
                                    if duration_weeks == 0:
                                        duration_weeks = 1 # Минимальная длительность для отображения
                                        
                                    tasks_data.append({
                                        "Phase": tab_name,
                                        "Description": str(desc_val).strip(),
                                        "Duration Weeks": duration_weeks
                                    })
                
                if not tasks_data and not milestones:
                    st.error("Не удалось найти данные на указанных вкладках. Проверьте названия вкладок и наличие колонки 'DESCRIPTION'.")
                    st.stop()

                # 3. Формирование нового Excel файла с Диаграммой Ганта
                out_wb = openpyxl.Workbook()
                out_ws = out_wb.active
                out_ws.title = "SMS Gantt Chart"
                
                # Определяем максимальную длительность для расчета количества колонок недель
                total_weeks = sum(task["Duration Weeks"] for task in tasks_data)
                max_weeks_display = max(20, min(total_weeks + 5, 52)) # Минимум 20, максимум 52 недели
                
                # --- Форматирование по требованию: ширина 4, шрифт 9 ---
                font_small = Font(size=9, name="Arial")
                thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), 
                                     top=Side(style='thin'), bottom=Side(style='thin'))
                
                # Заголовки
                headers = ["Phase", "Description", "Duration (Weeks)", "Start Date", "End Date"]
                for i, h in enumerate(headers, 1):
                    cell = out_ws.cell(row=1, column=i, value=h)
                    cell.font = Font(bold=True, size=10)
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal="center", vertical="center")
                
                # Заголовки недель + проверка на праздники РФ
                ru_holidays_set = get_ru_holidays(range(project_start_date.year, project_start_date.year + 2))
                current_week_start = project_start_date
                
                # Находим индекс первой колонки недели
                first_week_col = len(headers) + 1
                
                for w in range(1, max_weeks_display + 1):
                    col_letter = get_column_letter(first_week_col + w - 1)
                    
                    # Проверка на праздник в этой неделе
                    is_hol = is_holiday_week(current_week_start, ru_holidays_set)
                    week_label = f"W{w}\n{current_week_start.strftime('%d.%m')}"
                    if is_hol:
                        week_label += "\n🎉"
                    
                    cell = out_ws.cell(row=1, column=first_week_col + w - 1, value=week_label)
                    cell.font = font_small
                    cell.border = thin_border
                    cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                    
                    # Если праздник, красим фон заголовка в светло-красный
                    if is_hol:
                        cell.fill = PatternFill(start_color="FFCCCC", end_color="FFCCCC", fill_type="solid")
                    
                    # Устанавливаем ширину колонки = 4 (строго по ТЗ)
                    out_ws.column_dimensions[col_letter].width = 4
                    
                    current_week_start += timedelta(weeks=1)

                # Заполнение данными задач
                gantt_fill = PatternFill(start_color="4F81BD", end_color="4F81BD", fill_type="solid") # Синий цвет для Ганта
                current_date_cursor = project_start_date
                
                for row_idx, task in enumerate(tasks_data, start=2):
                    # Основные данные
                    out_ws.cell(row=row_idx, column=1, value=task["Phase"]).font = font_small
                    out_ws.cell(row=row_idx, column=2, value=task["Description"]).font = font_small
                    out_ws.cell(row=row_idx, column=3, value=task["Duration Weeks"]).font = font_small
                    
                    # Расчет дат (упрощенный: 1 неделя = 7 календарных дней, но с визуальной пометкой праздников)
                    # Если нужна строгая логика рабочих дней: duration_days = task["Duration Weeks"] * 5
                    # Но в Гантах по неделям обычно отображают календарные недели, помечая праздники.
                    start_dt = current_date_cursor
                    end_dt = start_dt + timedelta(weeks=task["Duration Weeks"])
                    
                    out_ws.cell(row=row_idx, column=4, value=start_dt.strftime("%d.%m.%Y")).font = font_small
                    out_ws.cell(row=row_idx, column=5, value=end_dt.strftime("%d.%m.%Y")).font = font_small
                    
                    # Рисуем бар Ганта
                    start_week_offset = 0 # Упрощенно: все задачи идут друг за другом или от общей даты. 
                    # Если нужно от общей даты начала проекта:
                    weeks_from_start = (start_dt - project_start_date).days // 7
                    
                    for w in range(task["Duration Weeks"]):
                        col_idx = first_week_col + weeks_from_start + w
                        if col_idx <= (first_week_col + max_weeks_display - 1):
                            cell = out_ws.cell(row=row_idx, column=col_idx, value="█") # Символ для наглядности в узкой колонке
                            cell.fill = gantt_fill
                            cell.font = Font(size=9, color="FFFFFF", name="Arial") # Белый шрифт на синем фоне
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                            cell.border = thin_border
                    
                    # Сдвигаем курсор для следующей задачи (последовательное выполнение)
                    # Если задачи параллельные, эту строку нужно убрать или модифицировать логику
                    current_date_cursor = end_dt

                # Применяем границы ко всем заполненным ячейкам
                for row in out_ws.iter_rows(min_row=1, max_row=len(tasks_data)+1, min_col=1, max_col=first_week_col+max_weeks_display-1):
                    for cell in row:
                        if not cell.border.left: # Если граница еще не установлена
                            cell.border = thin_border

                # Сохранение в буфер
                output = io.BytesIO()
                out_wb.save(output)
                output.seek(0)
                
                st.success("✅ График успешно сформирован!")
                st.dataframe(pd.DataFrame(tasks_data).head(10)) # Превью первых 10 строк
                
                st.download_button(
                    label="📥 Скачать SMS График (Excel)",
                    data=output,
                    file_name="SMS_Gantt_Chart_TOGF-ENG.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet"
                )
                
            except Exception as e:
                st.error(f"Произошла ошибка при обработке файла: {str(e)}")
                st.exception(e)

