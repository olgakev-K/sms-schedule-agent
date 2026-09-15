import datetime
import io
import openpyxl
from openpyxl.chart import BarChart, Reference
from openpyxl.styles import Alignment, Border, Font, PatternFill, Side
from openpyxl.utils import get_column_letter
import plotly.express as px
import pandas as pd
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
    (1, 8),
    (2, 23),
    (3, 8),
    (5, 1),
    (5, 9),
    (6, 12),
    (11, 4),
]


def is_working_day(date_val):
    if date_val.weekday() >= 5: # Суббота, Воскресенье
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


# --- 2. ИНТЕРФЕЙС STREAMLIT ---
st.set_page_config(
    page_title="SMS График запуска (TOGF-ENG)", page_icon="📊", layout="wide"
)

st.title("📊 SMS-график проекта запуска в серийное производство")
st.caption(
    "Сводка проектов и визуальная Диаграмма Ганта на одном листе (согласно consultant.ru)"
)

uploaded_file = st.file_uploader(
    "Загрузите Excel-файл шаблона PLANT_MASTER_SCHEDULE P25077.xlsx:",
    type=["xlsx"],
)

if uploaded_file:
    with st.spinner(
        "Извлечение данных, расчет дат и длительностей по AM, AN, AO, AP..."
    ):
        try:
            wb_src = openpyxl.load_workbook(uploaded_file, data_only=False)

            # 1. Сбор вех и определение даты старта
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

            # 2. Сбор задач из вкладок в строго заданной последовательности
            target_sheets = [
                ("Phase 2", "PMSPR TOGF-ENG-008-06 Phase2"),
                ("Phase 3", "PMSPR TOGF-ENG-008-06 Phase 3"),
                ("Phase 4;5", "PMSPR TOGF-ENG-008-06 Phase4;5"),
            ]

            tasks = []
            phase_summary = {}
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

                    # Расчет Duration weeks по цветным ячейкам (AM=39, AN=40, AO=41, AP=42...)
                    duration_weeks = 0
                    for c in range(39, 80):
                        cell = ws.cell(row=r, column=c)
                        fill = cell.fill
                        has_color = False
                        if fill and fill.fill_type is not None:
                            color_rgb = getattr(fill.start_color, "rgb", None)
                            if color_rgb and str(color_rgb) not in [
                                "00000000",
                                "FFFFFFFF",
                                "00FFFFFF",
                            ]:
                                has_color = True
                        if has_color or cell.value is not None:
                            duration_weeks += 1
                        elif duration_weeks > 0:
                            break

                    if duration_weeks == 0:
                        duration_weeks = 1

                    # Расчет даты окончания (игнорируем плановые даты исходника)
                    task_end = calculate_end_date(curr_start, duration_weeks)

                    tasks.append(
                        {
                            "Phase": phase_label,
                            "Task": str(desc).strip(),
                            "Duration": duration_weeks,
                            "Start": curr_start,
                            "End": task_end,
                        }
                    )

                    phase_duration += duration_weeks
                    curr_start = task_end + datetime.timedelta(days=1)

                if phase_duration > 0:
                    phase_summary[phase_label] = {
                        "Start": phase_start,
                        "End": tasks[-1]["End"],
                        "Duration": phase_duration,
                    }

            df_tasks = pd.DataFrame(tasks)

            # --- ЭКРАННОЕ ПРЕВЬЮ (STREAMLIT) ---
            st.subheader("📋 Сводный список проектов")
            df_phases = pd.DataFrame(
                [
                    {
                        "Проект / Фаза": k,
                        "Дата начала": v["Start"].strftime("%d.%m.%Y"),
                        "Дата окончания": v["End"].strftime("%d.%m.%Y"),
                        "Общая длит. (нед)": v["Duration"],
                    }
                    for k, v in phase_summary.items()
                ]
            )
            st.dataframe(df_phases, use_container_width=True)

            st.subheader("📊 Экранная диаграмма Ганта")
            fig = px.timeline(
                df_tasks,
                x_start="Start",
                x_end="End",
                y="Task",
                color="Phase",
                title="Календарный график",
            )
            fig.update_yaxes(autorange="reversed")
            fig.update_layout(height=450, margin=dict(l=10, r=10, t=30, b=10))
            st.plotly_chart(fig, use_container_width=True)

            # --- ГЕНЕРАЦИЯ ЕДИНОГО EXCEL-ЛИСТА (СПИСОК + ГАНТ + СТАТИСТИКА) ---
            wb_out = openpyxl.Workbook()
            ws_out = wb_out.active
            ws_out.title = "SMS Master Schedule"

            # Стили
            f_norm = Font(name="Arial", size=9)
            f_bold = Font(name="Arial", size=9, bold=True)
            f_title = Font(name="Arial", size=11, bold=True)

            fill_head = PatternFill(
                start_color="F2F2F2", end_color="F2F2F2", fill_type="solid"
            )
            fill_bar = PatternFill(
                start_color="1F497D", end_color="1F497D", fill_type="solid"
            )
            fill_hol = PatternFill(
                start_color="FFC7CE", end_color="FFC7CE", fill_type="solid"
            )

            b_thin = Border(
                left=Side(style="thin", color="D9D9D9"),
                right=Side(style="thin", color="D9D9D9"),
                top=Side(style="thin", color="D9D9D9"),
                bottom=Side(style="thin", color="D9D9D9"),
            )

            # 1. СПИСОК ПРОЕКТОВ (В верхней части этого же листа)
            ws_out.cell(
                row=1, column=1, value="СВОДНЫЙ СПИСОК ПРОЕКТОВ / ФАЗ"
            ).font = f_title

            p_headers = [
                "Проект / Фаза",
                "Дата начала",
                "Дата окончания",
                "Длит. (нед)",
            ]
            for c_i, h_text in enumerate(p_headers, 1):
                cell = ws_out.cell(row=3, column=c_i, value=h_text)
                cell.font = f_bold
                cell.fill = fill_head
                cell.border = b_thin
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )

            p_row = 4
            for p_name, p_info in phase_summary.items():
                ws_out.cell(row=p_row, column=1, value=p_name).font = f_norm
                ws_out.cell(
                    row=p_row,
                    column=2,
                    value=p_info["Start"].strftime("%d.%m.%Y"),
                ).font = f_norm
                ws_out.cell(
                    row=p_row,
                    column=3,
                    value=p_info["End"].strftime("%d.%m.%Y"),
                ).font = f_norm
                ws_out.cell(
                    row=p_row, column=4, value=p_info["Duration"]
                ).font = f_norm

                for c in range(1, 5):
                    ws_out.cell(row=p_row, column=c).border = b_thin
                    if c > 1:
                        ws_out.cell(row=p_row, column=c).alignment = Alignment(
                            horizontal="center"
                        )
                p_row += 1

            # 2. ДЕТАЛИЗИРОВАННАЯ ДИАГРАММА ГАНТА (Начиная ниже на этом же листе)
            gantt_start_row = p_row + 3
            ws_out.cell(
                row=gantt_start_row - 2,
                column=1,
                value="КАЛЕНДАРНЫЙ ГРАФИК ЗАДАЧ (ДИАГРАММА ГАНТА)",
            ).font = f_title

            # Отображение вех над шапкой диаграммы (строки N-1, N)
            col_m = 7
            for code, val in milestones.items():
                ws_out.cell(
                    row=gantt_start_row - 1, column=col_m, value=str(code)
                ).font = f_bold
                ws_out.cell(
                    row=gantt_start_row, column=col_m, value="◆"
                ).font = f_bold
                col_m += 1

            headers = [
                "№",
                "Фаза",
                "Описание действия (Description)",
                "Длит. (нед)",
                "Дата начала",
                "Дата окончания",
            ]
            for c_idx, h_text in enumerate(headers, 1):
                cell = ws_out.cell(
                    row=gantt_start_row, column=c_idx, value=h_text
                )
                cell.font = f_bold
                cell.fill = fill_head
                cell.alignment = Alignment(
                    horizontal="center", vertical="center"
                )

            # Отрисовка колонок недель W1..Wn
            total_weeks = sum(t["Duration"] for t in tasks)
            w_curr_dt = auto_start_date

            for w in range(1, total_weeks + 1):
                col_idx = 6 + w
                w_end = w_curr_dt + datetime.timedelta(days=6)
                is_hol = has_holiday_in_range(w_curr_dt, w_end)

                w_title = f"W{w} 🎉" if is_hol else f"W{w}"
                c_w = ws_out.cell(
                    row=gantt_start_row, column=col_idx, value=w_title
                )
                c_w.font = f_bold
                c_w.fill = fill_hol if is_hol else fill_head
                c_w.alignment = Alignment(
                    horizontal="center", vertical="center"
                )

                # Ширина столбца недели = 4
                ws_out.column_dimensions[get_column_letter(col_idx)].width = 4
                w_curr_dt += datetime.timedelta(days=7)

            # Вывод всех строк с задачами и графическими барами Ганта
            w_pointer = 1
            for i, t in enumerate(tasks, 1):
                r_idx = gantt_start_row + i
                ws_out.cell(row=r_idx, column=1, value=i)
                ws_out.cell(row=r_idx, column=2, value=t["Phase"])
                ws_out.cell(row=r_idx, column=3, value=t["Task"])
                ws_out.cell(row=r_idx, column=4, value=t["Duration"])
                ws_out.cell(
                    row=r_idx, column=5, value=t["Start"].strftime("%d.%m.%Y")
                )
                ws_out.cell(
                    row=r_idx, column=6, value=t["End"].strftime("%d.%m.%Y")
                )

                # Отрисовка баров Ганта прямо на сетке
                for dw in range(t["Duration"]):
                    gc = ws_out.cell(
                        row=r_idx, column=6 + w_pointer + dw, value="█"
                    )
                    gc.fill = fill_bar
                    gc.font = Font(name="Arial", size=9, color="1F497D")

                w_pointer += t["Duration"]

                for col in range(1, 7 + total_weeks):
                    cell = ws_out.cell(row=r_idx, column=col)
                    cell.font = f_norm
                    cell.border = b_thin
                    if col in [1, 4, 5, 6]:
                        cell.alignment = Alignment(horizontal="center")

            # 3. ДОБАВЛЕНИЕ ВСТРОЕННОЙ ГРАФИЧЕСКОЙ ДИАГРАММЫ EXCEL (Chart Object)
            chart = BarChart()
            chart.type = "bar"
            chart.style = 10
            chart.title = "Сводная длительность задач"
            chart.y_axis.title = "Задачи"
            chart.x_axis.title = "Длительность (в неделях)"

            data_ref = Reference(
                ws_out,
                min_col=4,
                min_row=gantt_start_row,
                max_row=gantt_start_row + len(tasks),
            )
            cats_ref = Reference(
                ws_out,
                min_col=3,
                min_row=gantt_start_row + 1,
                max_row=gantt_start_row + len(tasks),
            )

            chart.add_data(data_ref, titles_from_data=True)
            chart.set_categories(cats_ref)
            chart.height = 12
            chart.width = 18

            # Размещение графического объекта правее таблицы фаз
            ws_out.add_chart(chart, "F1")

            # Настройка ширины колонок
            ws_out.column_dimensions["A"].width = 5
            ws_out.column_dimensions["B"].width = 15
            ws_out.column_dimensions["C"].width = 45
            ws_out.column_dimensions["D"].width = 14
            ws_out.column_dimensions["E"].width = 14
            ws_out.column_dimensions["F"].width = 14

            output = io.BytesIO()
            wb_out.save(output)
            output.seek(0)

            st.download_button(
                label="📥 Скачать итоговый Excel-файл (Сводка + Гант)",
                data=output,
                file_name="SMS_Master_Schedule_Combined.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
            )

        except Exception as e:
            st.error(f"❌ Ошибка обработки файла: {str(e)}")
