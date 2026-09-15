import datetime
import openpyxl
from openpyxl.styles import Alignment, Font, PatternFill

# 1. Проверка даты старта
start_date_str = input_user_date  # Передается пользователем (например, "01.10.2026")
if not start_date_str:
    raise ValueError(
        "Пожалуйста, укажите дату начала проекта в формате ДД.ММ.ГГГГ."
    )

start_date = datetime.datetime.strptime(start_date_str, "%d.%m.%Y").date()

# 2. Извлечение вех из START PROJECT TOGF-ENG-007-02
wb = openpyxl.load_workbook("PLANT_MASTER_SCHEDULE P25077.xlsx", data_only=True)
ws_milestones = wb["START PROJECT TOGF-ENG-007-02"]

milestones = []
for row in range(30, 36):
    name = ws_milestones[f"B{row}"].value  # Буквенное обозначение / имя
    date_val = ws_milestones[f"C{row}"].value  # Дата или значение
    if name:
        milestones.append({"code": str(name).strip(), "date_val": date_val})

# 3. Чтение фаз и задач
sheets_to_process = [
    ("Phase 2", "PMSPR TOGF-ENG-008-06 Phase2"),
    ("Phase 3", "PMSPR TOGF-ENG-008-06 Phase 3"),
    ("Phase 4;5", "PMSPR TOGF-ENG-008-06 Phase4;5"),
]

tasks = []
current_start = start_date

for phase_name, sheet_name in sheets_to_process:
    if sheet_name not in wb.sheetnames:
        continue
    ws = wb[sheet_name]

    # Находим колонку DESCRIPTION
    desc_col = None
    for col in range(1, 20):
        val = str(ws.cell(row=1, column=col).value or "").upper()
        if "DESCRIPTION" in val or "ОПИСАНИЕ" in val or "ДЕЙСТВИЕ" in val:
            desc_col = col
            break
    if not desc_col:
        desc_col = 2  # по умолчанию колонка B

    # Перебор строк с задачами
    for r in range(2, ws.max_row + 1):
        desc = ws.cell(row=r, column=desc_col).value
        if not desc:
            continue

        # Подсчет длительности (сканирование от AM (колонка 39) вправо)
        duration_weeks = 0
        for c in range(39, 60):  # AM, AN, AO, AP...
            cell = ws.cell(row=r, column=c)
            # Проверка наличия заливки или данных
            fill_color = cell.fill.start_color.rgb if cell.fill else None
            has_fill = (
                fill_color is not None
                and fill_color != "00000000"
                and fill_color != "FFFFFFFF"
            )
            has_data = cell.value is not None

            if has_fill or has_data:
                duration_weeks += 1
            elif duration_weeks > 0:
                # Прерываем подсчет, если цепочка смежных ячеек закончилась
                break

        if duration_weeks == 0:
            duration_weeks = 1  # Минимальная длительность

        task_end = current_start + datetime.timedelta(days=duration_weeks * 7)

        tasks.append({
            "phase": phase_name,
            "description": str(desc).strip(),
            "duration": duration_weeks,
            "start": current_start,
            "end": task_end,
        })

        current_start = task_end  # Следующая задача начинается по окончании текущей

# 4. Формирование итогового отчета в Excel (openpyxl)
# - Создание колонок W1..Wn
# - Установка ширины столбцов недель = 4
# - Установка шрифта 9pt для всех ячеек
# - Подсветка праздников РФ (#FFC7CE) в шапке Wn
# - Отрисовка вех (код над строкой 9 и символ в строке 10) по образцу BK9:BK10

