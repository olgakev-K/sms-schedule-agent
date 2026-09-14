import streamlit as st
import pandas as pd
import numpy as np
import datetime
import requests
from bs4 import BeautifulSoup
import holidays

st.set_page_config(page_title="SMS Schedule Agent", layout="wide")

st.title("🤖 ИИ-Агент: Генератор SMS-графика проекта")
st.write("Автоматический расчет актуального календарного плана с учетом праздников РФ (КонсультантПлюс)")

# Функция для получения праздников РФ
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
        
    ru_holidays = holidays.RU(years=[datetime.datetime.now().year, datetime.datetime.now().year + 1])
    for d in ru_holidays.keys():
        holiday_dates.add(np.datetime64(d, 'D'))
        
    return list(holiday_dates)

# Поиск автоматической даты старта в Excel
def extract_start_date_from_excel(xls, phases):
    for sheet in phases:
        if sheet in xls.sheet_names:
            df = pd.read_excel(xls, sheet_name=sheet, header=None)
            
            # Поиск столбца START DATE / Baseline Start / Start
            start_col = None
            header_row = None
            for r in range(min(15, df.shape[0])):
                for c in range(df.shape[1]):
                    val = str(df.iloc[r, c]).strip().upper()
                    if val in ['START DATE', 'START', 'BASELINE START', 'PLAN START', 'ДАТА НАЧАЛА']:
                        start_col = c
                        header_row = r
                        break
                if start_col is not None:
                    break
            
            if start_col is not None and header_row is not None:
                # Извлекаем первую валидную дату из этого столбца
                for r in range(header_row + 1, df.shape[0]):
                    val = df.iloc[r, start_col]
                    if pd.notna(val):
                        dt = pd.to_datetime(val, errors='coerce')
                        if pd.notna(dt):
                            return dt.date()
    return None

# Загрузка файла Excel пользователем
uploaded_file = st.file_uploader("Загрузите шаблон Excel (PLANT_MASTER_SCHEDULE P25077.xlsx)", type=["xlsx"])

phases = [
    'PMSPR TOGF-ENG-008-06 Phase2',
    'PMSPR TOGF-ENG-008-06 Phase 3',
    'PMSPR TOGF-ENG-008-06 Phase4;5'
]

detected_start_date = None
if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        detected_start_date = extract_start_date_from_excel(xls, phases)
    except Exception:
        pass

default_date = detected_start_date if detected_start_date else datetime.date.today()

if detected_start_date:
    st.info(f"🎯 Дата начала проекта автоматически определена из файла Excel: **{detected_start_date.strftime('%d.%m.%Y')}**")

start_date = st.date_input("Дата старта проекта (определена из файла или выберите вручную):", default_date)

if uploaded_file and st.button("Сформировать SMS-график"):
    with st.spinner("Агент обрабатывает данные и производственный календарь..."):
        try:
            xls = pd.ExcelFile(uploaded_file)
            
            tasks = []
            for sheet in phases:
                if sheet in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet, header=None)
                    desc_col = None
                    for r in range(min(15, df.shape[0])):
                        for c in range(df.shape[1]):
                            if str(df.iloc[r, c]).strip() == 'DESCRIPTION':
                                desc_col = c
                                break
                        if desc_col is not None:
                            break
                    if desc_col is not None:
                        for val in df.iloc[10:, desc_col].dropna():
                            tasks.append({'Phase': sheet, 'DESCRIPTION': str(val).strip()})

            df_tasks = pd.DataFrame(tasks)
            
            holiday_dates = get_rf_holidays()
            bus_cal = np.busday_calendar(weekmask='1111100', holidays=holiday_dates)
            
            current_date = np.datetime64(start_date, 'D')
            schedule = []
            
            for idx, row in df_tasks.iterrows():
                if not np.is_busday(current_date, busdaycal=bus_cal):
                    current_date = np.busday_offset(current_date, 0, roll='forward', busdaycal=bus_cal)
                
                t_start = current_date
                t_end = np.busday_offset(t_start, 0, roll='forward', busdaycal=bus_cal)
                
                schedule.append({
                    '№': idx + 1,
                    'Фаза проекта': row['Phase'],
                    'DESCRIPTION': row['DESCRIPTION'],
                    'Дата начала (расчет)': pd.to_datetime(str(t_start)).strftime('%d.%m.%Y'),
                    'Дата окончания (расчет)': pd.to_datetime(str(t_end)).strftime('%d.%m.%Y')
                })
                
                current_date = np.busday_offset(t_end, 1, roll='forward', busdaycal=bus_cal)
            
            res_df = pd.DataFrame(schedule)
            st.success(f"График успешно сформирован! Извлечено задач: {len(res_df)}")
            
            st.dataframe(res_df, use_container_width=True)
            
            excel_out = "SMS_Schedule_Result.xlsx"
            with pd.ExcelWriter(excel_out, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False, sheet_name='SMS Schedule')
            
            with open(excel_out, "rb") as f:
                st.download_button("📥 Скачать итоговый Excel", f, file_name="SMS_Schedule_Result.xlsx")
                
        except Exception as e:
            st.error(f"Ошибка при обработке файла: {e}")
