import datetime
import io
import openpyxl
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import streamlit as st

# --- ПРОИЗВОДСТВЕННЫЙ КАЛЕНДАРЬ РФ ---
# Список государственных праздников РФ (месяц, день) согласно consultant.ru
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
    """Проверяет, является ли день рабочим (исключает СБ, ВС и праздники РФ)"""
    if date_val.weekday() >= 5: # 5 = Суббота, 6 = Воскресенье
        return False
    if (date_val.month, date_val.day) in RU_HOLIDAYS:
        return False
    return True


def add_working_days(start_dt, num_weeks):
    """Рассчитывает дату окончания с учетом рабочих дней и праздников РФ"""
    target_working_days = num_weeks * 5 # 1 неделя длительности = 5 рабочих дней
    current_dt = start_dt
    added_days = 0

    while added_days < target_working_days:
        current_dt += datetime.timedelta(days=1)
        if is_working_day(current_dt):
            added_days += 1

    return current_dt


def has_holiday_in_range(start_dt, end_dt):
    """Проверяет наличие праздничного дня в периоде"""
    curr = start_dt
    while curr <= end_dt:
        if (curr.month, curr.day) in RU_HOLIDAYS:
            return True
        curr += datetime.timedelta(days=1)
    return False


# --- STREAMLIT ИНТЕРФЕЙС ---
st.set_page_config(
    page_title="SMS График запуска (TOGF-ENG)", page_icon="📊", layout="wide"
)

st.title("📊 Генератор SMS-графика проекта запуска в серийное производство")
st.caption(
    "Автоматический расчет плана с учетом производственного календаря РФ (consultant.ru)"
)

uploaded_file = st.file_uploader(
    "Загрузите файл шаблона PLANT_MASTER_SCHEDULE P25077.xlsx:", type=["xlsx"]
)

if uploaded_file:
    with st.spinner("Анализ файла и извлечение данных..."):
        try:
            wb_src = openpyxl.load_workbook(uploaded_file, data_only=True)

            # --- 1. АВТООПРЕДЕЛЕНИЕ ДАТЫ СТАРТА И СБОР ВЕХ ---
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

                        # Пробуем извлечь дату старта из ячейки вехи
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
                    f"⚠️ В ячейках C30:C35 дата старта не найдена. Автоматически установлена дата: **{auto_start_date.strftime('%d.%m.%Y')}**"
                )
            else:
                st.success(
                    f"📅 Дата начала проекта автоматически определена из вех: **{auto_start_date.strftime('%d.%m.%Y')}**"
                )

            # --- 2. СБОР ЗАДАЧ ИЗ ВИНКЛЮЧЕННЫХ ВКЛАДОК ---
            target_sheets = [
                ("Phase 2", "PMSPR TOGF-ENG-008-06 Phase2"),
                ("Phase 3", "PMSPR TOGF-ENG-008-06 Phase 3"),
                ("Phase 4;5", "PMSPR TOGF-ENG-008-06 Phase4;5"),
            ]

            tasks = []
            curr_start = auto_start_date

            for phase_name, sheet_name in target_sheets:
                if sheet_name not in wb_src.sheetnames:
                    continue

                ws = wb_src[sheet_name]

                # Находим колонку DESCRIPTION
                desc_col = None
                for c in range(1, 25):
                    val_h = str(ws.cell(row=1, column=c).value or "").upper()
                    if "DESCRIPTION" in val_h or "ОПИСАНИЕ" in val_h:
                        desc_col = c
                        break
                if not desc_col:
                    desc_col = 2 # По умолчанию B

                # Чтение строк с описанием
                for r in range(2, ws.max_row + 1):
                    desc = ws.cell(row=r, column=desc_col).value
                    if not desc or not str(desc).strip():
                        continue

                    # Игнорируем авто-колонки. Расчет длительности по заливке от AM (39)
                    duration_weeks = 0
                    for c in range(39, 70):
                        cell = ws.cell(row=r, column=c)
                        fill = cell.fill
                        has_fill = (
                            fill
                            and fill.fill_type is not None
                            and fill.start_color.rgb
                            not in ["00000000", "FFFFFFFF", None]
                        )
                        if has_fill or cell.value is not None:
                            duration_weeks += 1
                        elif duration_weeks > 0:
                            break

                    if duration_weeks == 0:
                        duration_weeks = 1 # Мин. длительность 1 неделя

                    # Расчет даты окончания с учетом производственного календаря
                    task_end = add_working_days(curr_start, duration_weeks)

                    tasks.append(
                        {
                            "phase": phase_name,
                            "desc": str(desc).strip(),
                            "duration": duration_weeks,
                            "start": curr_start,
                            "end": task_end,
                        }
                    )

                    # Следующая задача начинается со следующего дня за окончанием предыдущей
                    curr_start = task_end + datetime.timedelta(days=1)

            # --- 3. ГЕНЕРАЦИЯ ИТОГОВОГО EXCEL ---
            if st.button("🚀 Сформировать и скачать SMS-график"):
                wb_out = openpyxl.Workbook()
                ws_out = wb_out.active
                ws_out.title = "SMS Schedule"

                # Стили оформления
                f_norm = Font(name="Arial", size=9)
                f_bold = Font(name="Arial", size=9, bold=True)
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

                # Шапка
                headers = [
                    "№",
                    "Фаза",
                    "Описание действия (Description)",
                    "Длит. (нед)",
                    "Дата начала",
                    "Дата окончания",
                ]
                for c_idx, h_text in enumerate(headers, 1):
                    cell = ws_out.cell(row=10, column=c_idx, value=h_text)
                    cell.font = f_bold
                    cell.fill = fill_head
                    cell.alignment = Alignment(
                        horizontal="center", vertical="center"
                    )

                # Отрисовка вех (по образцу строки 9-10)
                col_m = 7
                for code, val in milestones.items():
                    ws_out.cell(row=9, column=col_m, value=str(code)).font = (
                        f_bold
                    )
                    ws_out.cell(row=10, column=col_m, value="◆").font = f_bold
                    col_m += 1

                # Заполнение задач
                total_weeks = sum(t["duration"] for t in tasks)
                w_curr_dt = auto_start_date

                # Недели W1..Wn
                for w in range(1, total_weeks + 1):
                    col_idx = 6 + w
                    w_end = w_curr_dt + datetime.timedelta(days=6)
                    is_hol = has_holiday_in_range(w_curr_dt, w_end)

                    w_title = f"W{w} 🎉" if is_hol else f"W{w}"
                    c_w = ws_out.cell(row=10, column=col_idx, value=w_title)
                    c_w.font = f_bold
                    c_w.fill = fill_hol if is_hol else fill_head
                    c_w.alignment = Alignment(
                        horizontal="center", vertical="center"
                    )

                    # Ширина колонок недель = 4
                    ws_out.column_dimensions[
                        get_column_letter(col_idx)
                    ].width = 4
                    w_curr_dt += datetime.timedelta(days=7)

                # Строки с данными
                w_pointer = 1
                for i, t in enumerate(tasks, 1):
                    r_idx = 10 + i
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

                    # Диаграмма Ганта
                    for dw in range(t["duration"]):
                        gc = ws_out.cell(
                            row=r_idx,
                            column=6 + w_pointer + dw,
                            value="█",
                        )
                        gc.fill = fill_bar
                        gc.font = Font(name="Arial", size=9, color="1F497D")

                    w_pointer += t["duration"]

                    # Форматирование ячеек
                    for col in range(1, 7 + total_weeks):
                        cell = ws_out.cell(row=r_idx, column=col)
                        cell.font = f_norm
                        cell.border = b_thin
                        if col in [1, 4, 5, 6]:
                            cell.alignment = Alignment(horizontal="center")

                # Настройка ширины колонок
                ws_out.column_dimensions["A"].width = 5
                ws_out.column_dimensions["B"].width = 14
                ws_out.column_dimensions["C"].width = 48
                ws_out.column_dimensions["D"].width = 12
                ws_out.column_dimensions["E"].width = 14
                ws_out.column_dimensions["F"].width = 14

                output = io.BytesIO()
                wb_out.save(output)
                output.seek(0)

                st.download_button(
                    label="📥 Скачать готовый SMS-график (.xlsx)",
                    data=output,
                    file_name="SMS_Master_Schedule.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                )

        except Exception as e:
            st.error(f"❌ Ошибка обработки файла: {str(e)}")
