import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from datetime import datetime, timedelta
import holidays
import io

st.set_page_config(page_title="SMS Gantt Generator", page_icon="📊", layout="wide")
st.title("📊 Генератор SMS графика проекта (TOGF-ENG)")
st.markdown("""
**Инструкция:**
1. Загрузите файл шаблона `PLANT_MASTER_SCHEDULE P25077.xlsx`
2. Укажите дату начала проекта
3. Нажмите "Сформировать SMS График"
4. Скачайте готовый Excel-файл с диаграммой Ганта
""")

# ============================================================
# ФУНКЦИИ
# ============================================================

def get_ru_holidays(years_range):
    """Возвращает множество дат государственных праздников РФ"""
    try:
        return set(holidays.RU(years=years_range))
    except:
        # Fallback: основные праздники РФ
        ru_holidays = set()
        for year in years_range:
            ru_holidays.update([
                datetime(year, 1, 1).date(), datetime(year, 1, 2).date(),
                datetime(year, 1, 3).date(), datetime(year, 1, 4).date(),
                datetime(year, 1, 5).date(), datetime(year, 1, 6).date(),
                datetime(year, 1, 7).date(), datetime(year, 1, 8).date(),
                datetime(year, 2, 23).date(),
                datetime(year, 3, 8).date(),
                datetime(year, 5, 1).date(), datetime(year, 5, 9).date(),
                datetime(year, 6, 12).date(),
                datetime(year, 11, 4).date(),
            ])
        return ru_holidays


def count_colored_weeks(ws, row_idx, start_col="AM", end_col="AZ"):
    """
    Считает количество залитых цветом ячеек в строке.
    Каждая залитая ячейка = 1 неделя длительности.
    """
    start_col_idx = column_index_from_string(start_col)
    end_col_idx = column_index_from_string(end_col)
    count = 0

    for col_idx in range(start_col_idx, end_col_idx + 1):
        cell = ws.cell(row=row_idx, column=col_idx)
        fill = cell.fill

        # Проверяем наличие заливки
        has_fill = False
        if fill and fill.fgColor:
            rgb = fill.fgColor
            # rgb может быть строкой или объектом
            rgb_str = str(rgb) if rgb else ""
            # Исключаем прозрачный и белый
            if rgb_str not in ['00000000', '00FFFFFF', 'FFFFFF', 'FFFFFFFF', 'None', '']:
                has_fill = True

        if has_fill:
            count += 1

    return count


def find_data_start_row(ws, keyword="DESCRIPTION", max_rows=20):
    """Находит строку с заголовком DESCRIPTION"""
    for row in range(1, max_rows + 1):
        for col in range(1, min(ws.max_column + 1, 50)):
            cell = ws.cell(row=row, column=col)
            if cell.value and isinstance(cell.value, str):
                if keyword.lower() in cell.value.lower():
                    return row, col
    return None, None


def find_description_column(ws, max_rows=10):
    """Ищет колонку с DESCRIPTION"""
    for row in range(1, max_rows + 1):
        for col in range(1, min(ws.max_column + 1, 50)):
            cell = ws.cell(row=row, column=col)
            if cell.value and isinstance(cell.value, str):
                val_lower = cell.value.lower()
                if "description" in val_lower or "описание" in val_lower or "действие" in val_lower:
                    return col
    return None


# ============================================================
# ИНТЕРФЕЙС
# ============================================================

uploaded_file = st.file_uploader(
    "📁 Загрузите файл 'PLANT_MASTER_SCHEDULE P25077.xlsx'",
    type=["xlsx", "xls"]
)

project_start = st.date_input(
    "📅 Дата начала проекта:",
    value=datetime.now()
)

if st.button(" Сформировать SMS График", type="primary", disabled=not uploaded_file):

    if uploaded_file is None:
        st.warning("Пожалуйста, загрузите файл.")
        st.stop()

    with st.spinner(" Обработка файла..."):
        try:
            # Загружаем workbook
            wb = openpyxl.load_workbook(uploaded_file, data_only=True)

            # Показываем доступные вкладки для отладки
            st.info(f"📋 Найденные вкладки в файле: {wb.sheetnames}")

            # ------------------------------------------------
            # 1. ВЕХИ из "START PROJECT TOGF-ENG-007-02"
            # ------------------------------------------------
            milestones = []
            milestone_sheet = "START PROJECT TOGF-ENG-007-02"

            if milestone_sheet in wb.sheetnames:
                ws_m = wb[milestone_sheet]
                st.success(f"✅ Вкладка '{milestone_sheet}' найдена")

                # Ищем колонку с описанием вех
                desc_col_m = find_description_column(ws_m)
                if desc_col_m is None:
                    desc_col_m = 1  # По умолчанию первая колонка

                for row in range(2, min(ws_m.max_row + 1, 200)):
                    val = ws_m.cell(row=row, column=desc_col_m).value
                    if val and str(val).strip():
                        milestones.append(str(val).strip())

                st.write(f"📌 Найдено вех: {len(milestones)}")
            else:
                st.warning(f"️ Вкладка '{milestone_sheet}' не найдена")

            # ------------------------------------------------
            # 2. ЗАДАЧИ из Phase2, Phase 3, Phase4;5
            # ------------------------------------------------
            phase_sheets = [
                "PMSPR TOGF-ENG-008-06 Phase2",
                "PMSPR TOGF-ENG-008-06 Phase 3",
                "PMSPR TOGF-ENG-008-06 Phase4;5"
            ]

            all_tasks = []

            for sheet_name in phase_sheets:
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    st.success(f"✅ Вкладка '{sheet_name}' найдена (строк: {ws.max_row})")

                    # Ищем колонку DESCRIPTION
                    desc_col = find_description_column(ws)

                    if desc_col is None:
                        st.warning(f"⚠️ Колонка DESCRIPTION не найдена в '{sheet_name}'. Использую колонку A.")
                        desc_col = 1

                    # Ищем строку заголовка
                    header_row, _ = find_data_start_row(ws, "DESCRIPTION")
                    if header_row is None:
                        header_row = 1

                    # Читаем данные
                    for row in range(header_row + 1, min(ws.max_row + 1, 500)):
                        desc_val = ws.cell(row=row, column=desc_col).value

                        if desc_val and str(desc_val).strip():
                            # Считаем залитые ячейки от AM до AZ
                            duration = count_colored_weeks(ws, row, "AM", "AZ")

                            # Если не нашли заливку — ставим 1 неделю минимум
                            if duration == 0:
                                duration = 1

                            all_tasks.append({
                                "Phase": sheet_name,
                                "Description": str(desc_val).strip(),
                                "Duration Weeks": duration
                            })

                    st.write(f"   → Задач из '{sheet_name}': {len([t for t in all_tasks if t['Phase'] == sheet_name])}")
                else:
                    st.warning(f"️ Вкладка '{sheet_name}' не найдена")

            if not all_tasks:
                st.error("❌ Не удалось найти задачи ни в одной из вкладок. Проверьте названия вкладок в файле.")
                st.stop()

            st.success(f"✅ Всего задач для графика: {len(all_tasks)}")

            # ------------------------------------------------
            # 3. СОЗДАНИЕ EXCEL С ДИАГРАММОЙ ГАНТА
            # ------------------------------------------------
            out_wb = openpyxl.Workbook()
            out_ws = out_wb.active
            out_ws.title = "SMS Gantt Chart"

            # Стили
            font_9 = Font(size=9, name="Arial")
            font_9_bold = Font(size=9, name="Arial", bold=True)
            font_header = Font(size=10, name="Arial", bold=True, color="FFFFFF")
            thin_border = Border(
                left=Side(style='thin'),
                right=Side(style='thin'),
                top=Side(style='thin'),
                bottom=Side(style='thin')
            )
            gantt_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            gantt_fill_alt = PatternFill(start_color="70AD47", end_color="70AD47", fill_type="solid")
            header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
            holiday_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            milestone_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")

            # --- Заголовки основных колонок ---
            main_headers = ["№", "Phase", "Description", "Duration (Weeks)", "Start Date", "End Date"]
            for i, h in enumerate(main_headers, 1):
                cell = out_ws.cell(row=1, column=i, value=h)
                cell.font = font_header
                cell.fill = header_fill
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

            # Устанавливаем ширину основных колонок
            col_widths = [5, 25, 50, 15, 14, 14]
            for i, w in enumerate(col_widths, 1):
                out_ws.column_dimensions[get_column_letter(i)].width = w

            # --- Колонки недель ---
            ru_holidays_set = get_ru_holidays(range(project_start.year, project_start.year + 2))

            # Рассчитываем общее количество недель
            total_weeks_needed = sum(t["Duration Weeks"] for t in all_tasks) + 5
            max_weeks = min(max(total_weeks_needed, 20), 104)  # Максимум 2 года

            first_week_col = len(main_headers) + 1
            current_week_date = project_start

            for w in range(1, max_weeks + 1):
                col_letter = get_column_letter(first_week_col + w - 1)

                # Проверяем праздник
                is_holiday = False
                for day_offset in range(7):
                    check_date = current_week_date + timedelta(days=day_offset)
                    if check_date in ru_holidays_set:
                        is_holiday = True
                        break

                # Формируем заголовок
                week_label = f"W{w}\n{current_week_date.strftime('%d.%m')}"
                if is_holiday:
                    week_label += "\n🎉"

                cell = out_ws.cell(row=1, column=first_week_col + w - 1, value=week_label)
                cell.font = font_9_bold
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

                if is_holiday:
                    cell.fill = holiday_fill
                else:
                    cell.fill = header_fill

                # ШИРИНА КОЛОНКИ = 4 (по ТЗ)
                out_ws.column_dimensions[col_letter].width = 4

                current_week_date += timedelta(weeks=1)

            # --- Заполнение данными задач ---
            current_date = project_start

            for idx, task in enumerate(all_tasks, start=2):
                # Основные данные
                out_ws.cell(row=idx, column=1, value=idx - 1).font = font_9
                out_ws.cell(row=idx, column=2, value=task["Phase"]).font = font_9
                out_ws.cell(row=idx, column=3, value=task["Description"]).font = font_9
                out_ws.cell(row=idx, column=4, value=task["Duration Weeks"]).font = font_9

                start_dt = current_date
                end_dt = start_dt + timedelta(weeks=task["Duration Weeks"])

                out_ws.cell(row=idx, column=5, value=start_dt.strftime("%d.%m.%Y")).font = font_9
                out_ws.cell(row=idx, column=6, value=end_dt.strftime("%d.%m.%Y")).font = font_9

                # Рисуем бар Ганта
                weeks_from_start = 0
                fill_to_use = gantt_fill if idx % 2 == 0 else gantt_fill_alt

                for w in range(task["Duration Weeks"]):
                    col_idx = first_week_col + weeks_from_start + w
                    if col_idx <= first_week_col + max_weeks - 1:
                        cell = out_ws.cell(row=idx, column=col_idx)
                        cell.value = "█"
                        cell.fill = fill_to_use
                        cell.font = Font(size=9, color="FFFFFF", name="Arial")
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.border = thin_border

                # Применяем границы к основным ячейкам
                for c in range(1, len(main_headers) + 1):
                    out_ws.cell(row=idx, column=c).border = thin_border
                    out_ws.cell(row=idx, column=c).alignment = Alignment(vertical="center", wrap_text=True)

                # Сдвигаем дату для следующей задачи (последовательное выполнение)
                current_date = end_dt

            # --- Добавляем вехи в конец ---
            if milestones:
                milestone_row = len(all_tasks) + 3
                out_ws.cell(row=milestone_row, column=1, value="").border = thin_border
                out_ws.cell(row=milestone_row, column=2, value="ВЕХИ ПРОЕКТА").font = Font(size=10, bold=True, name="Arial")
                out_ws.cell(row=milestone_row, column=2).border = thin_border

                for i, m in enumerate(milestones):
                    r = milestone_row + 1 + i
                    out_ws.cell(row=r, column=1, value=i + 1).font = font_9
                    out_ws.cell(row=r, column=2, value="MILESTONE").font = font_9
                    out_ws.cell(row=r, column=3, value=m).font = font_9_bold
                    out_ws.cell(row=r, column=3).fill = milestone_fill
                    for c in range(1, len(main_headers) + 1):
                        out_ws.cell(row=r, column=c).border = thin_border

            # Заморозка панелей
            out_ws.freeze_panes = "C2"

            # Сохранение
            output = io.BytesIO()
            out_wb.save(output)
            output.seek(0)

            st.success("✅ График успешно сформирован!")

            # Превью
            st.subheader("📋 Предпросмотр задач:")
            df_preview = pd.DataFrame(all_tasks)
            st.dataframe(df_preview.head(20), use_container_width=True)

            # Кнопка скачивания
            st.download_button(
                label="📥 Скачать SMS График (Excel)",
                data=output,
                file_name=f"SMS_Gantt_Chart_{project_start.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary"
            )

        except Exception as e:
            st.error(f"❌ Ошибка: {str(e)}")
            import traceback
            st.code(traceback.format_exc())
