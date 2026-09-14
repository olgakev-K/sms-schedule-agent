import streamlit as st
import pandas as pd
import datetime
import requests
from bs4 import BeautifulSoup
import holidays

st.set_page_config(page_title="SMS Schedule Agent", layout="wide")

st.title("🤖 ИИ-Агент: Генератор SMS-графика проекта")
st.write("Автоматический расчет календарного плана на основе файла Excel и производственного календаря РФ")

# Функция получения праздников РФ
@st.cache_data
def get_rf_holidays():
    url = "https://www.consultant.ru/law/ref/calendar/proizvodstvennye/"
    headers = {"User-Agent": "Mozilla/5.0"}
    holiday_dates = set()
    
    try:
        response = requests.get(url, headers=headers, timeout=5)
        if response.status_code == 200:
            soup = BeautifulSoup(response.text, 'html.parser')
            for td in soup.find_all('td', class_=['holiday', 'work_short']):
                pass
    except Exception:
        pass
        
    current_year = datetime.datetime.now().year
    ru_holidays = holidays.RU(years=[current_year - 2, current_year - 1, current_year, current_year + 1, current_year + 2, current_year + 3])
    for d in ru_holidays.keys():
        holiday_dates.add(d)
        
    return holiday_dates

# Функция проверки, является ли день рабочим (пн-пт и не праздник РФ)
def is_business_day(date_val, holiday_dates):
    # 5 = Суббота, 6 = Воскресенье
    if date_val.weekday() >= 5:
        return False
    if date_val in holiday_dates:
        return False
    return True

# Переход к следующему рабочему дню
def get_next_business_day(date_val, holiday_dates):
    cur = date_val
    while not is_business_day(cur, holiday_dates):
        cur += datetime.timedelta(days=1)
    return cur

# Поиск названия проекта и даты начала в Excel
def extract_project_info(xls, phases):
    project_name = "Не определен"
    start_date = None
    
    # 1. Поиск названия проекта
    for sheet in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet, header=None)
        for r in range(min(15, df.shape[0])):
            for c in range(df.shape[1]):
                val = str(df.iloc[r, c]).strip().upper()
                if 'ПРОЕКТ' in val or 'PROJECT' in val:
                    if c + 1 < df.shape[1] and pd.notna(df.iloc[r, c + 1]):
                        project_name = str(df.iloc[r, c + 1]).strip()
                    elif c + 2 < df.shape[1] and pd.notna(df.iloc[r, c + 2]):
                        project_name = str(df.iloc[r, c + 2]).strip()
                    break
            if project_name != "Не определен":
                break
        if project_name != "Не определен":
            break

    # 2. Точечный поиск даты старта в колонках
    target_headers = ['PLANNED START DATE', 'PLANNED START', 'ДАТА / DATE', 'START DATE', 'DATE']
    
    for sheet in phases:
        if sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            
            for r in range(min(25, df.shape[0])):
                for c in range(df.shape[1]):
                    val = str(df.iloc[r, c]).strip().upper()
                    
                    if any(hdr in val for hdr in target_headers):
                        for r_val in range(r + 1, df.shape[0]):
                            cell_val = df.iloc[r_val, c]
                            if pd.notna(cell_val) and str(cell_val).strip().upper() not in ['N/A', 'NONE', 'CLOSED']:
                                dt = pd.to_datetime(cell_val, errors='coerce')
                                if pd.notna(dt):
                                    start_date = dt.date()
                                    return project_name, start_date
                                    
    return project_name, start_date

uploaded_file = st.file_uploader("Загрузите шаблон Excel (PLANT_MASTER_SCHEDULE P25077.xlsx)", type=["xlsx"])

phases = [
    'PMSPR TOGF-ENG-008-06 Phase2',
    'PMSPR TOGF-ENG-008-06 Phase 3',
    'PMSPR TOGF-ENG-008-06 Phase4;5'
]

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        project_name, project_start_date = extract_project_info(xls, phases)
        
        st.info(f"📌 **Название проекта:** {project_name}")
        
        if project_start_date:
            st.success(f"📅 **Дата начала проекта (извлечена из PLANNED START DATE):** {project_start_date.strftime('%d.%m.%Y')}")
            
            if st.button("🚀 Сформировать автоматический SMS-график"):
                with st.spinner("Агент формирует график с учетом праздников РФ..."):
                    tasks = []
                    
                    for sheet in phases:
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
                                    tasks.append({'Phase': sheet, 'DESCRIPTION': str(val).strip()})

                    df_tasks = pd.DataFrame(tasks)
                    holiday_dates = get_rf_holidays()
                    
                    current_date = project_start_date
                    schedule = []
                    
                    for idx, row in df_tasks.iterrows():
                        current_date = get_next_business_day(current_date, holiday_dates)
                        
                        t_start = current_date
                        t_end = t_start # Однодневный шаг на задачу
                        
                        schedule.append({
                            '№': idx + 1,
                            'Проект': project_name,
                            'Фаза проекта': row['Phase'],
                            'DESCRIPTION': row['DESCRIPTION'],
                            'Дата начала (расчет)': t_start.strftime('%d.%m.%Y'),
                            'Дата окончания (расчет)': t_end.strftime('%d.%m.%Y')
                        })
                        
                        current_date = get_next_business_day(current_date + datetime.timedelta(days=1), holiday_dates)
                    
                    res_df = pd.DataFrame(schedule)
                    st.success(f"График успешно сформирован! Извлечено задач: {len(res_df)}")
                    st.dataframe(res_df, use_container_width=True)
                    
                    excel_out = "SMS_Schedule_Result.xlsx"
                    with pd.ExcelWriter(excel_out, engine='openpyxl') as writer:
                        res_df.to_excel(writer, index=False, sheet_name='SMS Schedule')
                    
                    with open(excel_out, "rb") as f:
                        st.download_button("📥 Скачать итоговый Excel", f, file_name="SMS_Schedule_Result.xlsx")
        else:
            st.error("❌ Не удалось автоматически определить 'PLANNED START DATE' или 'Дата / Date' в файле. Проверьте заполнение таблицы.")
            
    except Exception as e:
        st.error(f"Ошибка при обработке файла: {e}")
else:
    st.info("ℹ️ Для запуска автоматического расчета загрузите Excel-файл.")
