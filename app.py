import datetime
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# --- 1. ПРОИЗВОДСТВЕННЫЙ КАЛЕНДАРЬ РФ (по данным consultant.ru) ---
RU_HOLIDAYS = [
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


def is_working_day(date_val):
    if date_val.weekday() >= 5: # СБ, ВС
        return False
    if (date_val.month, date_val.day) in RU_HOLIDAYS:
        return False
    return True


def calculate_end_date(start_dt, duration_weeks):
    target_working_days = duration_weeks * 5
    current_dt = start_dt
    added_days = 0

    while added_days < target_working_days:
        current_dt += datetime.timedelta(days=1)
        if is_working_day(current_dt):
            added_days += 1

    return current_dt


def has_holiday_in_range(start_dt, end_dt):
    curr = start_dt
    while curr <= end_dt:
        if (curr.month, curr.day) in RU_HOLIDAYS:
            return True
        curr += datetime.timedelta(days=1)
    return False


# --- 2. STREAMLIT ИНТЕРФЕЙС ---
st.set_page_config(
    page_title="SMS График запуска (TOGF-ENG)", page_icon="📊", layout="wide"
)

st.title("📊 Генератор SMS-графика проекта запуска в серийное производство")
st.caption(
    "Источник норм производственного календаря РФ: [Consultant.ru](https://www.consultant.ru/law/ref/calendar/proizvodstvennye/)"
)

uploaded_file = st.file_uploader(
    "Загрузите файл шаблона PLANT_MASTER_SCHEDULE P25077.xlsx:", type=["xlsx"]
)

if uploaded_file:
    with st.spinner("Чтение данных и анализ файлов..."):
        try:
            wb_src = openpyxl.load_workbook(uploaded_file, data_only=False)

            # --- УСЛОВИЕ 1: Сбор вех ---
            milestones = {}
            auto_start_date = None

            if "START PROJECT TOGF-ENG-007-02" in wb_src.sheetnames:
                ws_m = wb_src["START PROJECT TOGF-ENG-007-02"]
                for r in range(30, 36):
                    code = ws_m[f"B{r}"].value
                    val = ws_m[f"C{r}"].value
                    if code:
                        code_str = str(code).strip()
                        milestones[code_str] = val

                        if not auto_start_date and isinstance(
                            val, (datetime.datetime, datetime.date)
                        ):
                            auto_start_date = (
                                val.date()
                                if isinstance(val, datetime.datetime)
                                else val
                            )

            if not auto_start_date:
                auto_start_date = datetime.date.today()
                st.warning(
                    f"⚠️ Дата старта не найдена. Установлена текущая дата: **{auto_start_date.strftime('%d.%m.%Y')}**"
                )
            else:
                st.success(
                    f"📅 Дата начала проекта определена автоматически: **{auto_start_date.strftime('%d.%m.%Y')}**"
                )

            # --- УСЛОВИЕ 2 & 5: Чтение задач и подрасчет по фазам/проектам ---
            target_sheets = [
                ("Phase 2", "PMSPR TOGF-ENG-008-06 Phase2"),
                ("Phase 3", "PMSPR TOGF-ENG-008-06 Phase 3"),
                ("Phase 4;5", "PMSPR TOGF-ENG-008-06 Phase4;5"),
            ]

            tasks = []
            phase_summary = (
                {}
            ) # Данные для сводного списка проектов на той же странице
            curr_start = auto_start_date

            for phase_label, sheet_name in target_sheets:
                if sheet_name not in wb_src.sheetnames:
                    continue

                ws = wb_src[sheet_name]

                desc_col = None
                for c in range(1, 25):
                    val_h = str(ws.cell(row=1, column=c).value or "").upper()
                    if "DESCRIPTION" in val_h or "ОПИСАНИЕ" in val_h:
                        desc_col = c
                        break
                if not desc_col:
                    desc_col = 2

                phase_start = curr_start
                phase_duration = 0

                for r in range(2, ws.max_row + 1):
                    desc = ws.cell(row=r, column=desc_col).value
                    if not desc or not str(desc).strip():
                        continue

                    # Расчет Duration weeks по цветной заливке AM, AN, AO, AP...
                    duration_weeks = 0
                    for c in range(39, 80):
                        cell = ws.cell(row=r, column=c)
                        fill = cell.fill

                        has_color_fill = False
                        if fill and fill.fill_type is not None:
                            color_rgb = getattr(fill.start_color, "rgb", None)
                            if color_rgb and str(color_rgb) not in [
                                "00000000",
                                "FFFFFFFF",
                                "00FFFFFF",
                            ]:
                                has_color_fill = True

                        if has_color_fill or cell.value is not None:
                            duration_weeks += 1
                        elif duration_weeks > 0:
                            break

                    if duration_weeks == 0:
                        duration_weeks = 1

                    task_end = calculate_end_date(curr_start, duration_weeks)

                    tasks.append(
                        {
                            "phase": phase_label,
                            "desc": str(desc).strip(),
                            "duration": duration_weeks,
                            "start": curr_start,
                            "end": task_end,
                        }
                    )

                    phase_duration += duration_weeks
                    curr_start = task_end + datetime.timedelta(days=1)

                if phase_duration > 0:
                    phase_summary[phase_label] = {
                        "start": phase_start,
                        "end": tasks[-1]["end"],
                        "duration": phase_duration,
                    }

            # --- 3. СОЗДАНИЕ ЕДИНОГО ЛИСТА В EXCEL ---
            if st.button("🚀 Сформировать и скачать SMS-график"):
                wb_out = openpyxl.Workbook()
                ws_out = wb_out.active
                ws_out.title = "SMS Master Schedule"

                # Стили
                f_norm = Font(name="Arial", size=9)
                f_bold = Font(name="Arial", size=9, bold=True)
                f_title = Font(name="Arial", size=11, bold=True)
                fill_bar = PatternFill(
                    start_color="1F497D",
                    end_color="1F497D",
                    fill_type="solid",
                )
                fill_hol = PatternFill(
                    start_color="FFC7CE",
                    end_color="FFC7CE",
                    fill_type="solid",
                )
                fill_head = PatternFill(
                    start_color="F2F2F2",
                    end_color="F2F2F2",
                    fill_type="solid",
                )
                b_thin = Border(
                    left=Side(style="thin", color="D9D9D9"),
                    right=Side(style="thin", color="D9D9D9"),
                    top=Side(style="thin", color="D9D9D9"),
                    bottom=Side(style="thin", color="D9D9D9"),
                )

                # --- 3.1. СВОДНЫЙ СПИСОК ПРОЕКТОВ (НА ЭТОЙ ЖЕ СТРАНИЦЕ, СТРОКИ 1-6) ---
                ws_out.cell(
                    row=1, column=1, value="СВОДНЫЙ СПИСОК ПРОЕКТОВ / ФАЗ"
                ).font = f_title

                p_headers = [
                    "Проект / Фаза",
                    "Дата начала",
                    "Дата окончания",
                    "Общая длит. (нед)",
                ]
                for c_i, h_text in enumerate(p_headers, 1):
                    cell = ws_out.cell(row=3, column=c_i, value=h_text)
                    cell.font = f_bold
                    cell.fill = fill_head
                    cell.border = b_thin
                    cell.alignment = Alignment(horizontal="center")

                p_row = 4
                for p_name, p_info in phase_summary.items():
                    ws_out.cell(row=p_row, column=1, value=p_name).font = f_norm
                    ws_out.cell(
                        row=p_row,
                        column=2,
                        value=p_info["start"].strftime("%d.%m.%Y"),
                    ).font = f_norm
                    ws_out.cell(
                        row=p_row,
                        column=3,
                        value=p_info["end"].strftime("%d.%m.%Y"),
                    ).font = f_norm
                    ws_out.cell(
                        row=p_row, column=4, value=p_info["duration"]
                    ).font = f_norm

                    for c in range(1, 5):
                        ws_out.cell(row=p_row, column=c).border = b_thin
                        if c > 1:
                            ws_out.cell(row=p_row, column=c).alignment = (
                                Alignment(horizontal="center")
                            )
                    p_row += 1

                # --- 3.2. ДИАГРАММА ГАНТА И ДЕТАЛИЗАЦИЯ (НА ЭТОЙ ЖЕ СТРАНИЦЕ, НАЧИНАЯ СО СТРОКИ 9) ---
                start_gantt_row = 10

                # Вехи проекта (строка 9)
                col_m = 7
                for code, val in milestones.items():
                    ws_out.cell(
                        row=start_gantt_row - 1, column=col_m, value=str(code)
                    ).font = f_bold
                    ws_out.cell(
                        row=start_gantt_row, column=col_m, value="◆"
                    ).font = f_bold
                    col_m += 1

                headers = [
                    "№",
                    "Фаза",
                    "Описание действия (Description)",
                    "Duration weeks",
                    "Дата начала",
                    "Дата окончания",
                ]
                for c_idx, h_text in enumerate(headers, 1):
                    cell = ws_out.cell(
                        row=start_gantt_row, column=c_idx, value=h_text
                    )
                    cell.font = f_bold
                    cell.fill = fill_head
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center"
                    )

                # Недели W1..Wn
                total_weeks = sum(t["duration"] for t in tasks)
                w_curr_dt = auto_start_date

                for w in range(1, total_weeks + 1):
                    col_idx = 6 + w
                    w_end = w_curr_dt + datetime.timedelta(days=6)
                    is_hol = has_holiday_in_range(w_curr_dt, w_end)

                    w_title = f"W{w} 🎉" if is_hol else f"W{w}"
                    c_w = ws_out.cell(
                        row=start_gantt_row, column=col_idx, value=w_title
                    )
                    c_w.font = f_bold
                    c_w.fill = fill_hol if is_hol else fill_head
                    c_w.alignment = Alignment(
                        horizontal="center", vertical="center"
                    )

                    ws_out.column_dimensions[
                        get_column_letter(col_idx)
                    ].width = 4
                    w_curr_dt += datetime.timedelta(days=7)

                # Строки задач и Ганта
                w_pointer = 1
                for i, t in enumerate(tasks, 1):
                    r_idx = start_gantt_row + i
                    ws_out.cell(row=r_idx, column=1, value=i)
                    ws_out.cell(row=r_idx, column=2, value=t["phase"])
                    ws_out.cell(row=r_idx, column=3, value=t["desc"])
                    ws_out.cell(row=r_idx, column=4, value=t["duration"])
                    ws_out.cell(
                        row=r_idx,
                        column=5,
                        value=t["start"].strftime("%d.%m.%Y"),
                    )
                    ws_out.cell(
                        row=r_idx,
                        column=6,
                        value=t["end"].strftime("%d.%m.%Y"),
                    )

                    # Бары диаграммы Ганта
                    for dw in range(t["duration"]):
                        gc = ws_out.cell(
                            row=r_idx,
                            column=6 + w_pointer + dw,
                            value="█",
                        )
                        gc.fill = fill_bar
                        gc.font = Font(name="Arial", size=9, color="1F497D")

                    w_pointer += t["duration"]

                    # Оформление ячеек
                    for col in range(1, 7 + total_weeks):
                        cell = ws_out.cell(row=r_idx, column=col)
                        cell.font = f_norm
                        cell.border = b_thin
                        if col in [1, 4, 5, 6]:
                            cell.alignment = Alignment(horizontal="center")

                # Размеры колонок
                ws_out.column_dimensions["A"].width = 5
                ws_out.column_dimensions["B"].width = 16
                ws_out.column_dimensions["C"].width = 50
                ws_out.column_dimensions["D"].width = 16
                ws_out.column_dimensions["E"].width = 14
                ws_out.column_dimensions["F"].width = 14

                output = io.BytesIO()
                wb_out.save(output)
                output.seek(0)

                st.download_button(
                    label="📥 Скачать готовый SMS-график (.xlsx)",
                    data=output,
                    file_name="SMS_Master_Schedule_P25077.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        except Exception as e:
            st.error(f"❌ Ошибка при обработке файла: {str(e)}")
