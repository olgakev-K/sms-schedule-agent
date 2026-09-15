import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles import Font, PatternFill, Alignment, Border, Side
from openpyxl.utils import get_column_letter, column_index_from_string
from datetime import datetime, timedelta
import holidays
import io
import traceback

st.set_page_config(page_title="ИИ-Агент: SMS Gantt Builder", page_icon="📊", layout="wide")

st.title("📊 ИИ-Агент: Генератор SMS-графика проекта (TOGF-ENG)")
st.markdown("""
**Автоматическое создание диаграммы Ганта из шаблона `PLANT_MASTER_SCHEDULE P25077`**
- 🎯 Вехи из вкладки "START PROJECT TOGF-ENG-007-02" интегрированы в диаграмму
- 📏 Длительность по цветовой заливке (от колонки AM)
- 📅 Учет производственного календаря РФ
- 🎨 Форматирование: ширина столбцов Ганта = 4, шрифт = 9
""")

# ============================================================
# ФУНКЦИИ
# ============================================================

def get_ru_holidays(years_range):
    try:
        return set(holidays.RU(years=years_range))
    except Exception:
        ru_holidays = set()
        for year in years_range:
            ru_holidays.update([
                datetime(year, 1, 1).date(), datetime(year, 1, 2).date(), datetime(year, 1, 3).date(),
                datetime(year, 1, 4).date(), datetime(year, 1, 5).date(), datetime(year, 1, 6).date(),
                datetime(year, 1, 7).date(), datetime(year, 1, 8).date(), datetime(year, 2, 23).date(),
                datetime(year, 3, 8).date(), datetime(year, 5, 1).date(), datetime(year, 5, 9).date(),
                datetime(year, 6, 12).date(), datetime(year, 11, 4).date(),
            ])
        return ru_holidays

def is_holiday_week(start_date, ru_holidays_set):
    for i in range(7):
        check_date = start_date + timedelta(days=i)
        if check_date in ru_holidays_set:
            return True
    return False

def find_description_column(ws, max_rows=10):
    for row in range(1, max_rows + 1):
        for col in range(1, min(ws.max_column + 1, 50)):
            cell = ws.cell(row=row, column=col)
            if cell.value and isinstance(cell.value, str):
                val_lower = str(cell.value).lower()
                if "description" in val_lower or "описание" in val_lower or "действие" in val_lower:
                    return col
    return 1

def extract_milestones_from_sheet(wb):
    """Извлекает вехи из вкладки START PROJECT TOGF-ENG-007-02"""
    sheet_name = "START PROJECT TOGF-ENG-007-02"
    milestones = []
    
    if sheet_name not in wb.sheetnames:
        return milestones
    
    ws = wb[sheet_name]
    milestone_col = None
    date_col = None
    
    # Ищем колонки с вехами и датами
    for row in range(1, 10):
        for col in range(1, 20):
            cell = ws.cell(row=row, column=col)
            if cell.value and isinstance(cell.value, str):
                val_lower = str(cell.value).lower()
                if any(kw in val_lower for kw in ["milestone", "веха", "этап", "gate", "review", "name", "наименование"]):
                    milestone_col = col
                if any(kw in val_lower for kw in ["date", "дата", "deadline", "срок", "target"]):
                    date_col = col
    
    if milestone_col is None:
        milestone_col = 1
    if date_col is None:
        date_col = milestone_col + 1
    
    for row in range(2, min(ws.max_row + 1, 100)):
        milestone_name = ws.cell(row=row, column=milestone_col).value
        milestone_date = ws.cell(row=row, column=date_col).value
        
        if milestone_name and str(milestone_name).strip():
            name_clean = str(milestone_name).strip()
            date_parsed = None
            
            if isinstance(milestone_date, datetime):
                date_parsed = milestone_date.date()
            elif milestone_date and isinstance(milestone_date, str):
                try:
                    date_parsed = datetime.strptime(milestone_date, "%d.%m.%Y").date()
                except:
                    try:
                        date_parsed = datetime.strptime(milestone_date, "%Y-%m-%d").date()
                    except:
                        pass
            
            milestones.append({"name": name_clean, "date": date_parsed, "row": row})
    
    return milestones

def count_colored_weeks(ws, row_idx, start_col="AM", end_col="ZZ"):
    start_col_idx = column_index_from_string(start_col)
    end_col_idx = column_index_from_string(end_col)
    count = 0

    for col_idx in range(start_col_idx, end_col_idx + 1):
        cell = ws.cell(row=row_idx, column=col_idx)
        fill = cell.fill
        
        has_fill = False
        if fill and fill.fgColor:
            rgb = fill.fgColor
            rgb_str = str(rgb).upper()
            if rgb_str not in ['00000000', '00FFFFFF', 'FFFFFF', 'FFFFFFFF', 'NONE', '']:
                if cell.value is not None and str(cell.value).strip() != "":
                    has_fill = True
        
        if has_fill:
            count += 1
            
    return count

# ============================================================
# ИНТЕРФЕЙС
# ============================================================

uploaded_file = st.file_uploader("📁 Загрузите файл 'PLANT_MASTER_SCHEDULE P25077.xlsx'", type=["xlsx"])

col1, col2 = st.columns(2)
with col1:
    date_mode = st.selectbox("📅 Режим расчета дат:", ["Параллельно (все от даты старта)", "Последовательно (друг за другом)"])
with col2:
    auto_date = None
    if uploaded_file:
        temp_wb = openpyxl.load_workbook(uploaded_file, data_only=True)
        # Пытаемся найти дату в файле
        if "START PROJECT TOGF-ENG-007-02" in temp_wb.sheetnames:
            ws_temp = temp_wb["START PROJECT TOGF-ENG-007-02"]
            for row in range(1, 30):
                for col in range(1, 15):
                    cell = ws_temp.cell(row=row, column=col)
                    if isinstance(cell.value, datetime):
                        header_cell = ws_temp.cell(row=row, column=col-1).value if col > 1 else ""
                        header_str = str(header_cell).lower() if header_cell else ""
                        if "start" in header_str or "начало" in header_str:
                            auto_date = cell.value.date()
                            break
    
    default_date = auto_date if auto_date else datetime.now().date()
    project_start = st.date_input("📍 Дата начала проекта:", value=default_date)
    
    if auto_date:
        st.success(f"✅ Авто-дата найдена: {auto_date}")

if st.button("🚀 Сформировать SMS График", type="primary", disabled=not uploaded_file, use_container_width=True):
    
    with st.spinner("🤖 ИИ-агент анализирует файл..."):
        try:
            wb = openpyxl.load_workbook(uploaded_file, data_only=True)
            
            # 1. Извлечение вех
            milestones = extract_milestones_from_sheet(wb)
            st.info(f"🎯 Найдено вех: {len(milestones)}")
            if milestones:
                for m in milestones[:5]:
                    date_str = m["date"].strftime("%d.%m.%Y") if m["date"] else "не указана"
                    st.write(f"  • {m['name']} — {date_str}")

            # 2. Извлечение задач
            phase_sheets = [
                "PMSPR TOGF-ENG-008-06 Phase2",
                "PMSPR TOGF-ENG-008-06 Phase 3",
                "PMSPR TOGF-ENG-008-06 Phase4;5"
            ]
            
            all_tasks = []
            for sheet_name in phase_sheets:
                if sheet_name in wb.sheetnames:
                    ws = wb[sheet_name]
                    desc_col = find_description_column(ws)
                    
                    for row in range(2, min(ws.max_row + 1, 1000)):
                        desc_val = ws.cell(row=row, column=desc_col).value
                        if desc_val and str(desc_val).strip():
                            duration = count_colored_weeks(ws, row, "AM", "ZZ")
                            if duration == 0:
                                duration = 1
                                
                            all_tasks.append({
                                "Phase": sheet_name,
                                "Description": str(desc_val).strip(),
                                "Duration Weeks": duration
                            })

            if not all_tasks:
                st.error("❌ Не удалось найти задачи.")
                st.stop()

            st.success(f"✅ Найдено задач: {len(all_tasks)}")

            # ============================================================
            # 3. ГЕНЕРАЦИЯ EXCEL
            # ============================================================
            out_wb = openpyxl.Workbook()
            out_ws = out_wb.active
            out_ws.title = "SMS Gantt Chart"

            font_9 = Font(size=9, name="Arial")
            font_9_bold = Font(size=9, name="Arial", bold=True)
            font_header = Font(size=9, name="Arial", bold=True, color="FFFFFF")
            thin_border = Border(left=Side(style='thin'), right=Side(style='thin'), top=Side(style='thin'), bottom=Side(style='thin'))
            
            gantt_fill = PatternFill(start_color="4472C4", end_color="4472C4", fill_type="solid")
            header_fill = PatternFill(start_color="2F5496", end_color="2F5496", fill_type="solid")
            holiday_fill = PatternFill(start_color="FFC7CE", end_color="FFC7CE", fill_type="solid")
            milestone_fill = PatternFill(start_color="FFC000", end_color="FFC000", fill_type="solid")
            milestone_marker_fill = PatternFill(start_color="FF0000", end_color="FF0000", fill_type="solid")

            main_headers = ["№", "Фаза", "Описание действия", "Длит. (нед)", "Начало", "Окончание"]
            col_widths = [5, 25, 50, 12, 12, 12]
            
            for i, h in enumerate(main_headers, 1):
                cell = out_ws.cell(row=1, column=i, value=h)
                cell.font = font_header
                cell.fill = header_fill
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                out_ws.column_dimensions[get_column_letter(i)].width = col_widths[i-1]

            ru_holidays_set = get_ru_holidays(range(project_start.year, project_start.year + 2))
            total_weeks_needed = sum(t["Duration Weeks"] for t in all_tasks) + 4
            max_weeks = min(max(total_weeks_needed, 20), 104)
            
            first_week_col = len(main_headers) + 1
            current_week_date = project_start

            for w in range(1, max_weeks + 1):
                col_letter = get_column_letter(first_week_col + w - 1)
                is_hol = is_holiday_week(current_week_date, ru_holidays_set)
                week_label = f"W{w}\n{current_week_date.strftime('%d.%m')}"
                if is_hol:
                    week_label += "\n🎉"
                
                cell = out_ws.cell(row=1, column=first_week_col + w - 1, value=week_label)
                cell.font = font_9_bold
                cell.border = thin_border
                cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
                cell.fill = holiday_fill if is_hol else header_fill
                out_ws.column_dimensions[col_letter].width = 4
                
                current_week_date += timedelta(weeks=1)

            # Заполнение задачами
            current_sequential_date = project_start
            
            for idx, task in enumerate(all_tasks, start=2):
                if date_mode == "Параллельно (все от даты старта)":
                    start_dt = project_start
                else:
                    start_dt = current_sequential_date
                    
                end_dt = start_dt + timedelta(weeks=task["Duration Weeks"])
                
                out_ws.cell(row=idx, column=1, value=idx - 1).font = font_9
                out_ws.cell(row=idx, column=2, value=task["Phase"]).font = font_9
                out_ws.cell(row=idx, column=3, value=task["Description"]).font = font_9
                out_ws.cell(row=idx, column=4, value=task["Duration Weeks"]).font = font_9
                out_ws.cell(row=idx, column=5, value=start_dt.strftime("%d.%m.%Y")).font = font_9
                out_ws.cell(row=idx, column=6, value=end_dt.strftime("%d.%m.%Y")).font = font_9
                
                weeks_from_start = (start_dt - project_start).days // 7
                
                for w in range(task["Duration Weeks"]):
                    col_idx = first_week_col + weeks_from_start + w
                    if col_idx <= first_week_col + max_weeks - 1:
                        cell = out_ws.cell(row=idx, column=col_idx)
                        cell.value = "█"
                        cell.fill = gantt_fill
                        cell.font = Font(size=9, color="FFFFFF", name="Arial")
                        cell.alignment = Alignment(horizontal="center", vertical="center")
                        cell.border = thin_border
                
                for c in range(1, len(main_headers) + 1):
                    out_ws.cell(row=idx, column=c).border = thin_border
                    out_ws.cell(row=idx, column=c).alignment = Alignment(vertical="center", wrap_text=True)
                
                if date_mode == "Последовательно (друг за другом)":
                    current_sequential_date = end_dt

            # ============================================================
            # 4. ДОБАВЛЕНИЕ ВЕХ В ДИАГРАММУ
            # ============================================================
            if milestones:
                m_start_row = len(all_tasks) + 3
                
                # Заголовок секции вех
                out_ws.cell(row=m_start_row, column=2, value="🎯 КЛЮЧЕВЫЕ ВЕХИ ПРОЕКТА").font = Font(size=10, bold=True, name="Arial", color="FFFFFF")
                out_ws.cell(row=m_start_row, column=2).fill = PatternFill(start_color="C00000", end_color="C00000", fill_type="solid")
                for c in range(1, len(main_headers) + 1):
                    out_ws.cell(row=m_start_row, column=c).border = thin_border
                
                # Каждая веха
                for i, m in enumerate(milestones):
                    r = m_start_row + 1 + i
                    
                    out_ws.cell(row=r, column=1, value=i + 1).font = font_9
                    out_ws.cell(row=r, column=2, value="ВЕХА").font = font_9_bold
                    out_ws.cell(row=r, column=3, value=m["name"]).font = font_9_bold
                    out_ws.cell(row=r, column=3).fill = milestone_fill
                    
                    if m["date"]:
                        out_ws.cell(row=r, column=5, value=m["date"].strftime("%d.%m.%Y")).font = font_9
                        
                        weeks_from_start = (m["date"] - project_start).days // 7
                        if 0 <= weeks_from_start < max_weeks:
                            col_idx = first_week_col + weeks_from_start
                            cell = out_ws.cell(row=r, column=col_idx)
                            cell.value = "▼"
                            cell.fill = milestone_marker_fill
                            cell.font = Font(size=9, color="FFFFFF", name="Arial", bold=True)
                            cell.alignment = Alignment(horizontal="center", vertical="center")
                            cell.border = thin_border
                    
                    for c in range(1, len(main_headers) + 1):
                        out_ws.cell(row=r, column=c).border = thin_border

            out_ws.freeze_panes = "C2"

            output = io.BytesIO()
            out_wb.save(output)
            output.seek(0)

            st.success("✅ График успешно сформирован!")
            
            st.subheader("📋 Предпросмотр задач:")
            df_preview = pd.DataFrame(all_tasks)
            st.dataframe(df_preview.head(15), use_container_width=True)

            st.download_button(
                label="📥 Скачать SMS График (Excel)",
                data=output,
                file_name=f"SMS_Gantt_Chart_{project_start.strftime('%Y%m%d')}.xlsx",
                mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                type="primary",
                use_container_width=True
            )

        except Exception as e:
            st.error(f"❌ Ошибка: {str(e)}")
            st.code(traceback.format_exc())
