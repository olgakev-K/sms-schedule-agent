# -*- coding: utf-8 -*-
"""
SMS график проекта запуска в серийное производство P25077
Источник: PLANT_MASTER_SCHEDULE P25077.xlsx
"""

import openpyxl
from openpyxl.utils import get_column_letter
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.chart import BarChart, Reference
import pandas as pd
from datetime import datetime, timedelta
import numpy as np

# ============================================================
# КОНФИГУРАЦИЯ
# ============================================================
SOURCE_FILE = "PLANT_MASTER_SCHEDULE P25077.xlsx"
OUTPUT_FILE = "SMS_P25077_Result.xlsx"

SHEET_MILESTONES = "START PROJECT TOGF-ENG-007-02"
SHEET_PHASES = [
    "PMSPR TOGF-ENG-008-06 Phase2",
    "PMSPR TOGF-ENG-008-06 Phase 3",
    "PMSPR TOGF-ENG-008-06 Phase4;5",
]

# Колонки для подсчёта Duration weeks (по заливке)
DURATION_COLUMNS = ["AM", "AN", "AQ", "AP"]

# Колонка с описанием действия
DESCRIPTION_COLUMN = "DESCRIPTION"

# Диапазон строк с вехами
MILESTONE_ROWS = range(30, 36)   # 30..35
MILESTONE_NAME_COL = "B"
MILESTONE_DATE_COL = "C"

# Игнорируемые колонки
IGNORED_COLUMNS = {
    "PLANNED START DATE",
    "PLANNED START WEEK (automatic)",
    "PLANNED END DATE (automatic)",
    "PLANNED END WEEK (automatic)",
    "ACTUAL END DATE STATUS (automatic)",
}

# Производственный календарь РФ (пример для 2024–2026)
# Загрузи актуальный список с consultant.ru и обнови здесь
RU_HOLIDAYS = {
    2024: [
        "2024-01-01","2024-01-02","2024-01-03","2024-01-04","2024-01-05",
        "2024-01-06","2024-01-07","2024-01-08","2024-02-23","2024-03-08",
        "2024-04-29","2024-04-30","2024-05-01","2024-05-09","2024-05-10",
        "2024-06-12","2024-11-04","2024-12-30","2024-12-31",
    ],
    2025: [
        "2025-01-01","2025-01-02","2025-01-03","2025-01-06","2025-01-07",
        "2025-01-08","2025-02-24","2025-03-10","2025-05-01","2025-05-02",
        "2025-05-08","2025-05-09","2025-06-12","2025-06-13","2025-11-03",
        "2025-11-04","2025-12-31",
    ],
    2026: [
        "2026-01-01","2026-01-02","2026-01-05","2026-01-06","2026-01-07",
        "2026-01-08","2026-02-23","2026-03-09","2026-05-01","2026-05-11",
        "2026-06-12","2026-11-04",
    ],
}


# ============================================================
# РАБОТА С ПРОИЗВОДСТВЕННЫМ КАЛЕНДАРЁМ
# ============================================================
def build_holiday_set(years):
    """Собирает множество дат-праздников для указанных лет."""
    holidays = set()
    for y in years:
        for d in RU_HOLIDAYS.get(y, []):
            holidays.add(datetime.strptime(d, "%Y-%m-%d").date())
    return holidays


def is_working_day(d, holidays):
    """Проверяет, является ли дата рабочим днём."""
    if d.weekday() >= 5:          # сб, вс
        return False
    if d in holidays:
        return False
    return True


def add_working_days(start_date, working_days, holidays):
    """Прибавляет N рабочих дней к дате."""
    current = start_date
    added = 0
    while added < working_days:
        current += timedelta(days=1)
        if is_working_day(current, holidays):
            added += 1
    return current


def weeks_to_working_days(weeks):
    """1 неделя = 5 рабочих дней."""
    return weeks * 5


# ============================================================
# ЧТЕНИЕ ФАЙЛА
# ============================================================
def read_milestones(wb):
    """Читает вехи из вкладки START PROJECT."""
    ws = wb[SHEET_MILESTONES]
    milestones = []
    for row in MILESTONE_ROWS:
        name = ws[f"{MILESTONE_NAME_COL}{row}"].value
        date_val = ws[f"{MILESTONE_DATE_COL}{row}"].value
        if name:
            milestones.append({
                "name": str(name).strip(),
                "date": date_val,
                "row": row,
            })
    return milestones


def count_filled_cells(ws, row, columns):
    """Считает количество залитых цветом ячеек в указанных столбцах."""
    filled = 0
    for col_letter in columns:
        cell = ws[f"{col_letter}{row}"]
        fill = cell.fill
        # Любая заливка, кроме "нет заливки"
        if fill is not None and fill.fill_type not in (None, "none"):
            # Проверяем, что цвет не белый и не прозрачный
            fg = fill.fgColor
            if fg is not None and fg.rgb not in (None, "00000000", "FFFFFFFF"):
                filled += 1
    return filled


def find_description_column(ws):
    """Ищет колонку DESCRIPTION в шапке (первые 20 строк)."""
    for row in range(1, 21):
        for col in range(1, 40):
            val = ws.cell(row=row, column=col).value
            if val and str(val).strip().upper() == DESCRIPTION_COLUMN.upper():
                return row, col
    return None, None


def read_phase_tasks(wb, sheet_name):
    """Читает задачи из вкладки фазы."""
    ws = wb[sheet_name]
    header_row, desc_col = find_description_column(ws)
    if header_row is None:
        print(f"⚠ Колонка DESCRIPTION не найдена на вкладке {sheet_name}")
        return []

    tasks = []
    for row in range(header_row + 1, ws.max_row + 1):
        desc = ws.cell(row=row, column=desc_col).value
        if not desc or not str(desc).strip():
            continue

        duration = count_filled_cells(ws, row, DURATION_COLUMNS)
        if duration == 0:
            continue  # пропускаем строки без заливки

        tasks.append({
            "phase": sheet_name,
            "description": str(desc).strip(),
            "duration_weeks": duration,
            "source_row": row,
        })
    return tasks


# ============================================================
# ПОСТРОЕНИЕ ГРАФИКА
# ============================================================
def build_schedule(tasks, milestones, holidays, start_date):
    """Строит календарный план с учётом рабочих дней."""
    # Определяем год(ы) для календаря
    years = set()
    for m in milestones:
        if isinstance(m["date"], datetime):
            years.add(m["date"].year)
        elif isinstance(m["date"], str):
            try:
                years.add(datetime.strptime(m["date"], "%Y-%m-%d").year)
            except ValueError:
                pass
    if not years:
        years = {start_date.year, start_date.year + 1}
    holidays = holidays | build_holiday_set(years)

    current = start_date
    for t in tasks:
        wd = weeks_to_working_days(t["duration_weeks"])
        t_start = current
        # если старт — выходной, сдвигаем на следующий рабочий
        while not is_working_day(t_start, holidays):
            t_start += timedelta(days=1)
        t_end = add_working_days(t_start, wd - 1, holidays) if wd > 0 else t_start
        t["start"] = t_start
        t["end"] = t_end
        t["duration_days"] = wd
        current = t_end + timedelta(days=1)  # следующая задача стартует после
    return tasks


# ============================================================
# ЭКСПОРТ В EXCEL С ДИАГРАММОЙ ГАНТА
# ============================================================
def export_to_excel(tasks, milestones, output_file):
    """Экспорт: слева таблица, справа диаграмма Ганта на одном листе."""
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "SMS P25077"

    # --- Стили ---
    header_font = Font(bold=True, color="FFFFFF", size=11)
    header_fill = PatternFill("solid", fgColor="1F4E78")
    milestone_fill = PatternFill("solid", fgColor="C00000")
    milestone_font = Font(bold=True, color="FFFFFF")
    center = Alignment(horizontal="center", vertical="center", wrap_text=True)
    left = Alignment(horizontal="left", vertical="center", wrap_text=True)
    thin = Side(border_style="thin", color="808080")
    border = Border(left=thin, right=thin, top=thin, bottom=thin)

    # --- Заголовок ---
    ws["A1"] = "SMS график запуска в серийное производство P25077"
    ws["A1"].font = Font(bold=True, size=14, color="1F4E78")
    ws.merge_cells("A1:H1")

    # --- Шапка таблицы ---
    headers = ["№", "Фаза", "Описание действия", "Начало",
               "Окончание", "Длительность (нед.)", "Длительность (раб. дн.)"]
    for i, h in enumerate(headers, start=1):
        c = ws.cell(row=3, column=i, value=h)
        c.font = header_font
        c.fill = header_fill
        c.alignment = center
        c.border = border

    # --- Данные ---
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

    # --- Ширина колонок ---
    ws.column_dimensions["A"].width = 5
    ws.column_dimensions["B"].width = 32
    ws.column_dimensions["C"].width = 55
    ws.column_dimensions["D"].width = 13
    ws.column_dimensions["E"].width = 13
    ws.column_dimensions["F"].width = 12
    ws.column_dimensions["G"].width = 14

    # --- Диаграмма Ганта (ручная отрисовка) ---
    # Стартовая колонка диаграммы — I
    gantt_start_col = 9
    gantt_start_row = 3

    # Определяем диапазон дат
    if tasks:
        min_date = min(t["start"] for t in tasks)
        max_date = max(t["end"] for t in tasks)
    else:
        min_date = max_date = datetime.today().date()

    # Шаг — одна колонка = 1 рабочий день (упрощённо календарный)
    total_days = (max_date - min_date).days + 1
    # Ограничим ширину диаграммы
    if total_days > 200:
        step = (max_date - min_date).days // 200 + 1
    else:
        step = 1

    # Заголовок диаграммы
    ws.cell(row=gantt_start_row - 1, column=gantt_start_col,
            value="Диаграмма Ганта").font = Font(bold=True, size=12)

    # Отрисовка задач
    for i, t in enumerate(tasks, start=1):
        r = gantt_start_row + i
        ws.cell(row=r, column=gantt_start_col - 1, value=i).alignment = center

        # Определяем колонки начала и конца
        start_offset = (t["start"] - min_date).days // step
        end_offset = (t["end"] - min_date).days // step

        for offset in range(start_offset, end_offset + 1):
            col = gantt_start_col + offset
            cell = ws.cell(row=r, column=col)
            cell.fill = PatternFill("solid", fgColor="4472C4")
            cell.border = border

    # --- Вехи как вертикальные маркеры ---
    milestone_font_small = Font(bold=True, color="C00000", size=10)
    for m in milestones:
        m_date = m["date"]
        if isinstance(m_date, str):
            try:
                m_date = datetime.strptime(m_date, "%Y-%m-%d").date()
            except ValueError:
                continue
        elif isinstance(m_date, datetime):
            m_date = m_date.date()
        else:
            continue

        offset = (m_date - min_date).days // step
        col = gantt_start_col + offset
        if col < gantt_start_col:
            continue

        # Буквенное обозначение вехи над диаграммой
        label_cell = ws.cell(row=gantt_start_row - 1, column=col, value=m["name"])
        label_cell.font = milestone_font_small
        label_cell.alignment = Alignment(horizontal="center", text_rotation=90)

        # Вертикальная линия — заливаем столбец
        for r in range(gantt_start_row, gantt_start_row + len(tasks) + 1):
            cell = ws.cell(row=r, column=col)
            if cell.fill is None or cell.fill.fill_type in (None, "none"):
                cell.fill = PatternFill("solid", fgColor="FFC000")

    # --- Легенда вех ---
    legend_row = gantt_start_row + len(tasks) + 3
    ws.cell(row=legend_row, column=gantt_start_col,
            value="Вехи проекта:").font = Font(bold=True)
    for i, m in enumerate(milestones, start=1):
        r = legend_row + i
        c1 = ws.cell(row=r, column=gantt_start_col, value=m["name"])
        c1.font = milestone_font_small
        c2 = ws.cell(row=r, column=gantt_start_col + 1,
                     value=str(m["date"]) if m["date"] else "")
        c2.alignment = left

    wb.save(output_file)
    print(f"✅ Файл сохранён: {output_file}")


# ============================================================
# ГЛАВНАЯ ФУНКЦИЯ
# ============================================================
def main():
    print("📂 Чтение файла:", SOURCE_FILE)
    wb = openpyxl.load_workbook(SOURCE_FILE, data_only=True)

    # 1. Вехи
    milestones = read_milestones(wb)
    print(f"✅ Найдено вех: {len(milestones)}")
    for m in milestones:
        print(f"   - {m['name']}: {m['date']}")

    # 2. Задачи по фазам
    all_tasks = []
    for sheet in SHEET_PHASES:
        tasks = read_phase_tasks(wb, sheet)
        print(f"✅ Вкладка '{sheet}': задач — {len(tasks)}")
        all_tasks.extend(tasks)

    if not all_tasks:
        print("❌ Задачи не найдены. Проверь колонку DESCRIPTION и заливку.")
        return

    # 3. Старт проекта — первая дата вехи или сегодня
    start_date = datetime.today().date()
    for m in milestones:
        if isinstance(m["date"], datetime):
            start_date = m["date"].date()
            break
        if isinstance(m["date"], str):
            try:
                start_date = datetime.strptime(m["date"], "%Y-%m-%d").date()
                break
            except ValueError:
                pass

    # 4. Построение календарного плана
    holidays = set()
    tasks = build_schedule(all_tasks, milestones, holidays, start_date)

    # 5. Экспорт
    export_to_excel(tasks, milestones, OUTPUT_FILE)


if __name__ == "__main__":
    main()
