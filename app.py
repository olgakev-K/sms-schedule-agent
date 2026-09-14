import streamlit as st
import pandas as pd
import datetime
import requests
from bs4 import BeautifulSoup
import holidays

st.set_page_config(page_title="SMS Schedule Agent", layout="wide")

st.title("🤖 ИИ-Агент: Генератор SMS-графика проекта")
st.write("Расчет календарного плана на основе вех из 'START PROJECT TOGF-ENG-007-02' и производственного календаря РФ")

# Получение праздников РФ
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

def is_business_day(date_val, holiday_dates):
    if date_val.weekday() >= 5 or date_val in holiday_dates:
        return False
    return True

def get_next_business_day(date_val, holiday_dates):
    cur = date_val
    while not is_business_day(cur, holiday_dates):
        cur += datetime.timedelta(days=1)
    return cur

# Условие 1: Извлечение вехи/даты старта ИЗ ВКЛАДКИ "START PROJECT TOGF-ENG-007-02"
def extract_start_milestone(xls):
    sheet_name = "START PROJECT TOGF-ENG-007-02"
    if sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        # Сканируем ячейки на предмет первой валидной даты
        for r in range(df.shape[0]):
            for c in range(df.shape[1]):
                cell_val = df.iloc[r, c]
                if pd.notna(cell_val) and str(cell_val).strip().upper() not in ['N/A', 'NONE', 'CLOSED']:
                    dt = pd.to_datetime(cell_val, errors='coerce')
                    if pd.notna(dt) and dt.year > 2000:
                        return dt.date()
    return None

uploaded_file = st.file_uploader("Загрузите файл (PLANT_MASTER_SCHEDULE P25077.xlsx)", type=["xlsx"])

# Условие 2: Последовательность вкладок
target_phases = [
    "PMSPR TOGF-ENG-008-06 Phase2",
    "PMSPR TOGF-ENG-008-06 Phase 3",
    "PMSPR TOGF-ENG-008-06 Phase4;5"
]

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        
        # Проверяем Условие 1
        start_date = extract_start_milestone(xls)
        
        if start_date:
            st.success(f"📅 Дата вехи извлечена из листа 'START PROJECT TOGF-ENG-007-02': **{start_date.strftime('%d.%m.%Y')}**")
            
            if st.button("🚀 Сформировать автоматический SMS-график"):
                with st.spinner("Формирование графика по правилам..."):
                    tasks = []
                    
                    # Условие 2: Загрузка DESCRIPTION из указанных вкладок по порядку
                    for sheet in target_phases:
                        if sheet in xls.sheet_names:
                            df = pd.read_excel(xls, sheet_name=sheet, header=None)
                            desc_col = None
                            
                            # Поиск колонки DESCRIPTION
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
                        df_tasks = pd.DataFrame(tasks)
                        holiday_dates = get_rf_holidays()
                        
                        current_date = start_date
                        schedule = []
                        
                        # Расчет дат с учетом производственного календаря РФ
                        for idx, row in df_tasks.iterrows():
                            current_date = get_next_business_day(current_date, holiday_dates)
                            
                            t_start = current_date
                            t_end = t_start
                            
                            schedule.append({
                                '№': idx + 1,
                                'Фаза проекта': row['Phase'],
                                'DESCRIPTION': row['DESCRIPTION'],
                                'Дата начала (расчет)': t_start.strftime('%d.%m.%Y'),
                                'Дата окончания (расчет)': t_end.strftime('%d.%m.%Y')
                            })
                            
                            current_date = get_next_business_day(current_date + datetime.timedelta(days=1), holiday_dates)
                        
                        res_df = pd.DataFrame(schedule)
                        st.success(f"SMS-график успешно сформирован! Задач: {len(res_df)}")
                        st.dataframe(res_df, use_container_width=True)
                        
                        excel_out = "SMS_Schedule_Result.xlsx"
                        with pd.ExcelWriter(excel_out, engine='openpyxl') as writer:
                            res_df.to_excel(writer, index=False, sheet_name='SMS Schedule')
                        
                        with open(excel_out, "rb") as f:
                            st.download_button("📥 Скачать итоговый Excel", f, file_name="SMS_Schedule_Result.xlsx")
                    else:
                        st.error("❌ Не удалось извлечь задачи из столбцов DESCRIPTION на указанных фазовых листах.")
        else:
            st.error("❌ Не удалось найти дату старта на вкладке 'START PROJECT TOGF-ENG-007-02'. Проверьте наличие этой вкладки и заполненость вех.")
            
    except Exception as e:
        st.error(f"Ошибка при обработке файла: {e}")
else:
    st.info("ℹ️ Для запуска расчета загрузите Excel-файл.")
