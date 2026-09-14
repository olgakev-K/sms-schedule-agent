import streamlit as st
import pandas as pd
import datetime
import holidays
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="SMS Schedule Agent", layout="wide")

st.title("🤖 ИИ-Агент: Генератор понятного Excel-графика проекта")
st.write("Формирование наглядной диаграммы Ганта по неделям с удобным скачиванием Excel-файла")

@st.cache_data
def get_rf_holidays():
    current_year = datetime.datetime.now().year
    ru_holidays = holidays.RU(years=[current_year - 2, current_year - 1, current_year, current_year + 1, current_year + 2, current_year + 3])
    return set(ru_holidays.keys())

def is_business_day(date_val, holiday_dates):
    if date_val.weekday() >= 5 or date_val in holiday_dates:
        return False
    return True

def get_next_business_day(date_val, holiday_dates):
    cur = date_val
    while not is_business_day(cur, holiday_dates):
        cur += datetime.timedelta(days=1)
    return cur

def extract_start_milestone(xls):
    sheet_name = "START PROJECT TOGF-ENG-007-02"
    if sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                cell_val = df.iloc[r, c]
                if pd.notna(cell_val) and str(cell_val).strip().upper() not in ['N/A', 'NONE', 'CLOSED']:
                    dt = pd.to_datetime(cell_val, errors='coerce')
                    if pd.notna(dt) and dt.year > 2000:
                        return dt.date()
    return None

def create_weekly_gantt_excel(schedule_data):
    wb = openpyxl.Workbook()
    ws = wb.active
    ws.title = "График работ (Гант)"
    ws.views.sheetView[0].showGridLines = True

    # Определение диапазона недель
    min_date = min(item['Start'] for item in schedule_data)
    max_date = max(item['Finish'] for item in schedule_data)

    # Список уникальных недель (год, номер недели)
    weeks = []
    curr = min_date
    while curr <= max_date + datetime.timedelta(days=7):
        iso_year, iso_week, _ = curr.isocalendar()
        if (iso_year, iso_week) not in weeks:
            weeks.append((iso_year, iso_week))
        curr += datetime.timedelta(days=7)

    # Цветовая гамма
    navy_fill = PatternFill(start_color="1F4E78", end_color="1F4E78", fill_type="solid")
    bar_fill = PatternFill(start_color="2F5597", end_color="2F5597", fill_type="solid")
    light_gray_fill = PatternFill(start_color="F2F2F2", end_color="F2F2F2", fill_type="solid")
    
    font_header = Font(name="Calibri", size=10, bold=True, color="FFFFFF")
    font_bold = Font(name="Calibri", size=10, bold=True)
    font_regular = Font(name="Calibri", size=10)

    thin_border = Border(
        left=Side(style='thin', color='D9D9D9'),
        right=Side(style='thin', color='D9D9D9'),
        top=Side(style='thin', color='D9D9D9'),
        bottom=Side(style='thin', color='D9D9D9')
    )

    # 1. Заголовки таблицы (Левая часть)
    base_headers = ["№", "Фаза проекта", "Описание работы (DESCRIPTION)", "Дата начала", "Дата финиша"]
    for col_idx, h in enumerate(base_headers, 1):
        cell = ws.cell(row=1, column=col_idx, value=h)
        cell.fill = navy_fill
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)

    # 2. Шкала недель (Правая часть)
    for w_idx, (y, w) in enumerate(weeks, len(base_headers) + 1):
        cell = ws.cell(row=1, column=w_idx, value=f"Неделя {w}\n({y})")
        cell.fill = navy_fill
        cell.font = font_header
        cell.alignment = Alignment(horizontal="center", vertical="center", wrap_text=True)
        ws.column_dimensions[get_column_letter(w_idx)].width = 11

    ws.row_dimensions[1].height = 30

    # 3. Заполнение строк данных и ячеек Ганта
    for r_idx, item in enumerate(schedule_data, 2):
        ws.cell(row=r_idx, column=1, value=item['№']).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=2, value=item['Фаза проекта'])
        ws.cell(row=r_idx, column=3, value=item['DESCRIPTION'])
        ws.cell(row=r_idx, column=4, value=item['Start'].strftime("%d.%m.%Y")).alignment = Alignment(horizontal="center")
        ws.cell(row=r_idx, column=5, value=item['Finish'].strftime("%d.%m.%Y")).alignment = Alignment(horizontal="center")

        # Получаем неделю задачи
        task_year, task_week, _ = item['Start'].isocalendar()

        # Оформление левой таблицы
        for c in range(1, len(base_headers) + 1):
            cell = ws.cell(row=r_idx, column=c)
            cell.font = font_regular
            cell.border = thin_border

        # Подсветка полосы Ганта по неделям
        for w_idx, (y, w) in enumerate(weeks, len(base_headers) + 1):
            cell = ws.cell(row=r_idx, column=w_idx)
            cell.border = thin_border
            if (y, w) == (task_year, task_week):
                cell.fill = bar_fill # Выделяем активную неделю синим цветом
            else:
                if (r_idx % 2) == 0:
                    cell.fill = light_gray_fill

        ws.row_dimensions[r_idx].height = 20

    # Настройка ширины колонок таблицы
    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 28
    ws.column_dimensions['C'].width = 50
    ws.column_dimensions['D'].width = 14
    ws.column_dimensions['E'].width = 14

    filename = "SMS_Project_Gantt_Schedule.xlsx"
    wb.save(filename)
    return filename

# --- Интерфейс Streamlit ---
uploaded_file = st.file_uploader("Загрузите Excel-файл проекта (PLANT_MASTER_SCHEDULE P25077.xlsx)", type=["xlsx"])
target_phases = ["PMSPR TOGF-ENG-008-06 Phase2", "PMSPR TOGF-ENG-008-06 Phase 3", "PMSPR TOGF-ENG-008-06 Phase4;5"]

if uploaded_file:
    xls = pd.ExcelFile(uploaded_file)
    start_date = extract_start_milestone(xls)
    
    if start_date:
        st.success(f"📅 Начальная дата вехи найдена: **{start_date.strftime('%d.%m.%Y')}**")
        
        tasks = []
        for sheet in target_phases:
            if sheet in xls.sheet_names:
                df = pd.read_excel(xls, sheet_name=sheet, header=None)
                desc_col = None
                for r in range(min(25, df.shape[0])):
                    for c in range(df.shape[1]):
                        if str(df.iloc[r, c]).strip().upper() == 'DESCRIPTION':
                            desc_col = c
                            break
                    if desc_col is not None:
                        break
                if desc_col is not None:
                    for val in df.iloc[10:, desc_col].dropna():
                        val_str = str(val).strip()
                        if val_str and val_str.upper() != 'DESCRIPTION':
                            tasks.append({'Phase': sheet, 'DESCRIPTION': val_str})

        if tasks:
            holiday_dates = get_rf_holidays()
            current_date = start_date
            schedule_data = []
            
            for idx, row in enumerate(tasks):
                current_date = get_next_business_day(current_date, holiday_dates)
                schedule_data.append({
                    '№': idx + 1,
                    'Фаза проекта': row['Phase'],
                    'DESCRIPTION': row['DESCRIPTION'],
                    'Start': current_date,
                    'Finish': current_date
                })
                current_date = get_next_business_day(current_date + datetime.timedelta(days=1), holiday_dates)

            # Генерация Excel
            excel_filename = create_weekly_gantt_excel(schedule_data)
            
            st.markdown("---")
            st.subheader("📊 Готовый график сформирован")
            st.write(f"Успешно обработано задач: **{len(schedule_data)}**")
            
            # Таблица предварительного просмотра
            df_preview = pd.DataFrame(schedule_data)
            df_preview['Start'] = df_preview['Start'].apply(lambda x: x.strftime('%d.%m.%Y'))
            df_preview['Finish'] = df_preview['Finish'].apply(lambda x: x.strftime('%d.%m.%Y'))
            st.dataframe(df_preview[['№', 'Фаза проекта', 'DESCRIPTION', 'Start', 'Finish']], use_container_width=True, height=300)

            # Прямая кнопка скачивания Excel
            with open(excel_filename, "rb") as file:
                st.download_button(
                    label="📥 СКАЧАТЬ ГРАФИК ГАНТА В EXCEL (.XLSX)",
                    data=file,
                    file_name="SMS_Project_Gantt_Schedule.xlsx",
                    mime="application/vnd.openxmlformats-officedocument.spreadsheetml.sheet",
                    use_container_width=True
                )
        else:
            st.error("Задачи не найдены в листах фаз.")
