import datetime
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# --- ПРАЗДНИКИ РФ ---
# Список государственных праздников РФ (месяц, день)
RUSSIAN_HOLIDAYS_MMDD = [
    (1, 1),
    (1, 2),
    (1, 3),
    (1, 4),
    (1, 5),
    (1, 6),
    (1, 7),
    (1, 8), # Новогодние каникулы
    (2, 23), # День защитника Отечества
    (3, 8), # Международный женский день
    (5, 1), # Праздник Весны и Труда
    (5, 9), # День Победы
    (6, 12), # День России
    (11, 4), # День народного единства
]


def is_holiday_in_week(week_start_date):
    """Проверяет, выпадает ли праздник РФ на 7 дней начиная с week_start_date"""
    for i in range(7):
        current_day = week_start_date + datetime.timedelta(days=i)
        if (current_day.month, current_day.day) in RUSSIAN_HOLIDAYS_MMDD:
            return True
    return False


# --- ИНТЕРФЕЙС STREAMLIT ---
st.set_page_config(
    page_title="Генератор SMS-графика", page_icon="📊", layout="wide"
)

st.title("📊 Генератор SMS-графика проекта (TOGF-ENG)")
st.caption(
    "Автоматический расчет Диаграммы Ганта из шаблона PLANT_MASTER_SCHEDULE"
)

# 1. Ввод даты
date_input = st.text_input(
    "Укажите дату начала проекта (ДД.ММ.ГГГГ):",
    placeholder="Например, 01.10.2026",
)

# 2. Загрузка файла
uploaded_file = st.file_uploader(
    "Загрузите Excel-файл шаблона (PLANT_MASTER_SCHEDULE):", type=["xlsx"]
)

if uploaded_file and date_input:
    try:
        start_date = datetime.datetime.strptime(date_input, "%d.%m.%Y").date()
    except ValueError:
        st.error(
            "❌ Неверный формат даты! Пожалуйста, используйте формат ДД.ММ.ГГГГ (например, 01.10.2026)."
        )
        st.stop()

    if st.button("🚀 Сформировать диаграмму Ганта", type="primary"):
        with st.spinner("Обработка данных и сборка отчета..."):
            try:
                wb_src = openpyxl.load_workbook(
                    uploaded_file, data_only=False
                )

                # --- 1. СБОР ВЕХ ---
                milestones = {}
                if "START PROJECT TOGF-ENG-007-02" in wb_src.sheetnames:
                    ws_m = wb_src["START PROJECT TOGF-ENG-007-02"]
                    for r in range(30, 36):
                        m_code = ws_m[f"B{r}"].value
                        if m_code:
                            m_code_str = str(m_code).strip()
                            milestones[m_code_str] = {
                                "code": m_code_str,
                                "val": ws_m[f"C{r}"].value,
                            }

                # --- 2. СБОР ЗАДАЧ ---
                target_sheets = [
                    ("Phase 2", "PMSPR TOGF-ENG-008-06 Phase2"),
                    ("Phase 3", "PMSPR TOGF-ENG-008-06 Phase 3"),
                    ("Phase 4;5", "PMSPR TOGF-ENG-008-06 Phase4;5"),
                ]

                tasks = []
                current_start = start_date

                for phase_label, sheet_name in target_sheets:
                    if sheet_name not in wb_src.sheetnames:
                        continue
                    ws = wb_src[sheet_name]

                    # Поиск колонки DESCRIPTION
                    desc_col = None
                    for c in range(1, 25):
                        val = str(ws.cell(row=1, column=c).value or "").upper()
                        if any(
                            k in val
                            for k in ["DESCRIPTION", "ОПИСАНИЕ", "ДЕЙСТВИЕ"]
                        ):
                            desc_col = c
                            break
                    if not desc_col:
                        desc_col = 2 # B по умолчанию

                    # Сканирование строк
                    for r in range(2, ws.max_row + 1):
                        desc = ws.cell(row=r, column=desc_col).value
                        if not desc or not str(desc).strip():
                            continue

                        # Расчет длительности по ячейкам AM (39) и далее
                        duration_weeks = 0
                        for c in range(39, 70):
                            cell = ws.cell(row=r, column=c)
                            fill = cell.fill
                            has_fill = (
                                fill
                                and fill.fill_type is not None
                                and fill.start_color.rgb not in ["00000000", "FFFFFFFF", None]
                            )
                            has_val = cell.value is not None

                            if has_fill or has_val:
                                duration_weeks += 1
                            elif duration_weeks > 0:
                                break

                        if duration_weeks == 0:
                            duration_weeks = 1

                        task_end = current_start + datetime.timedelta(
                            days=duration_weeks * 7
                        )

                        tasks.append(
                            {
                                "phase": phase_label,
                                "desc": str(desc).strip(),
                                "duration": duration_weeks,
                                "start": current_start,
                                "end": task_end,
                            }
                        )

                        current_start = task_end

                if not tasks:
                    st.error(
                        "❌ В указанных вкладках не найдены задачи для обработки."
                    )
                    st.stop()

                # --- 3. СОЗДАНИЕ ИТОГОВОГО ФАЙЛА EXCEL ---
                wb_out = openpyxl.Workbook()
                ws_out = wb_out.active
                ws_out.title = "Gantt Schedule"

                # Стили
                font_9 = Font(name="Arial", size=9)
                font_9_bold = Font(name="Arial", size=9, bold=True)
                fill_gantt = PatternFill(
                    start_color="1F497D",
                    end_color="1F497D",
                    fill_type="solid",
                ) # Темно-синий бар
                fill_holiday = PatternFill(
                    start_color="FFC7CE",
                    end_color="FFC7CE",
                    fill_type="solid",
                ) # Светло-красный
                fill_header = PatternFill(
                    start_color="F2F2F2",
                    end_color="F2F2F2",
                    fill_type="solid",
                )
                thin_border = Border(
                    left=Side(style="thin", color="D9D9D9"),
                    right=Side(style="thin", color="D9D9D9"),
                    top=Side(style="thin", color="D9D9D9"),
                    bottom=Side(style="thin", color="D9D9D9"),
                )

                # Заголовки таблицы
                headers = [
                    "№",
                    "Фаза (Phase)",
                    "Описание действия (Description)",
                    "Длит. (нед)",
                    "Дата начала",
                    "Дата окончания",
                ]
                total_weeks = sum(t["duration"] for t in tasks)

                # Шапка
                ws_out.row_dimensions[10].height = 20
                for c_idx, h_text in enumerate(headers, 1):
                    cell = ws_out.cell(row=10, column=c_idx, value=h_text)
                    cell.font = font_9_bold
                    cell.fill = fill_header
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center"
                    )

                # Заголовки недель (W1..Wn)
                week_starts = []
                w_start = start_date
                for w in range(1, total_weeks + 1):
                    col_idx = 6 + w
                    week_starts.append(w_start)

                    # Проверка праздников
                    is_hol = is_holiday_in_week(w_start)
                    w_header = f"W{w} 🎉" if is_hol else f"W{w}"

                    # Строки вех (9-10)
                    cell_w = ws_out.cell(row=10, column=col_idx, value=w_header)
                    cell_w.font = font_9_bold
                    cell_w.alignment = Alignment(
                        horizontal="center", vertical="center"
                    )
                    if is_hol:
                        cell_w.fill = fill_holiday
                    else:
                        cell_w.fill = fill_header

                    # Ширина колонок недель = 4
                    col_letter = get_column_letter(col_idx)
                    ws_out.column_dimensions[col_letter].width = 4

                    w_start += datetime.timedelta(days=7)

                # Отображение вех в строке 9 (по образцу BK9:BK10)
                ws_out.row_dimensions[9].height = 18
                for m_code, m_data in milestones.items():
                    # Примерное размещение первой вехи на старт
                    ws_out.cell(row=9, column=7, value=m_code).font = (
                        font_9_bold
                    )
                    ws_out.cell(row=10, column=7, value="◆").font = font_9_bold

                # Заполнение строк с задачами
                current_w_idx = 1
                for idx, t in enumerate(tasks, 1):
                    row_idx = 10 + idx
                    ws_out.cell(row=row_idx, column=1, value=idx)
                    ws_out.cell(row=row_idx, column=2, value=t["phase"])
                    ws_out.cell(row=row_idx, column=3, value=t["desc"])
                    ws_out.cell(row=row_idx, column=4, value=t["duration"])
                    ws_out.cell(
                        row=row_idx,
                        column=5,
                        value=t["start"].strftime("%d.%m.%Y"),
                    )
                    ws_out.cell(
                        row=row_idx,
                        column=6,
                        value=t["end"].strftime("%d.%m.%Y"),
                    )

                    # Отрисовка бара Ганта
                    for dw in range(t["duration"]):
                        g_col = 6 + current_w_idx + dw
                        cell_g = ws_out.cell(
                            row=row_idx, column=g_col, value="█"
                        )
                        cell_g.fill = fill_gantt
                        cell_g.font = Font(name="Arial", size=9, color="1F497D")

                    current_w_idx += t["duration"]

                    # Форматирование текста
                    for col in range(1, 7 + total_weeks):
                        cell = ws_out.cell(row=row_idx, column=col)
                        if col > 6:
                            cell.alignment = Alignment(horizontal="center")
                        cell.font = font_9
                        cell.border = thin_border

                # Настройка ширины основных колонок
                ws_out.column_dimensions["A"].width = 5
                ws_out.column_dimensions["B"].width = 15
                ws_out.column_dimensions["C"].width = 45
                ws_out.column_dimensions["D"].width = 12
                ws_out.column_dimensions["E"].width = 14
                ws_out.column_dimensions["F"].width = 14

                # Выгрузка в байты для скачивания
                output = io.BytesIO()
                wb_out.save(output)
                output.seek(0)

                st.success("✅ График успешно сформирован!")
                st.download_button(
                    label="📥 Скачать Диаграмму Ганта (Excel)",
                    data=output,
                    file_name="Gantt_Master_Schedule.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

            except Exception as e:
                st.error(f"❌ Ошибка при обработке файла: {str(e)}")

elif not date_input and uploaded_file:
    st.info("👈 Укажите дату начала проекта для запуска генерации.")

