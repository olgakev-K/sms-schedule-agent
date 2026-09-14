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

# Загрузка файла Excel пользователем
uploaded_file = st.file_uploader("Загрузите шаблон Excel (PLANT_MASTER_SCHEDULE P25077.xlsx)", type=["xlsx"])
start_date = st.date_input("Выберите дату старта проекта:", datetime.date.today())

if uploaded_file and st.button("Сформировать SMS-график"):
    with st.spinner("Агент обрабатывает данные и производственный календарь..."):
        try:
            xls = pd.ExcelFile(uploaded_file)
            
            phases = [
                'PMSPR TOGF-ENG-008-06 Phase2',
                'PMSPR TOGF-ENG-008-06 Phase 3',
                'PMSPR TOGF-ENG-008-06 Phase4;5'
            ]
            
            tasks = []
            for sheet in phases:
                if sheet in xls.sheet_names:
                    df = pd.read_excel(xls, sheet_name=sheet, header=None)
                    desc_col = None
                    for r in range(min(12, df.shape[0])):
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
                       excel_out = "SMS_Schedule_Result.xlsx"
            with pd.ExcelWriter(excel_out, engine='openpyxl') as writer:
                res_df.to_excel(writer, index=False, sheet_name='SMS Schedule')
            
            with open(excel_out, "rb") as f:
                st.download_button("📥 Скачать итоговый Excel", f, file_name="SMS_Schedule_Result.xlsx")
                
        except Exception as e:
            st.error(f"Ошибка при обработке файла: {e}")
