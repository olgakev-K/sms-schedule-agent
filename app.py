import streamlit as st
import pandas as pd
import datetime
import requests
from bs4 import BeautifulSoup
import holidays
import plotly.express as px
import plotly.graph_objects as go

st.set_page_config(page_title="SMS Schedule Agent", layout="wide")

st.title("🤖 ИИ-Агент: Генератор SMS-графика проекта")
st.write("Автоматический расчет календарного плана с визуализацией в виде Диаграммы Ганта")

# 4. Получение праздников РФ по производственному календарю
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
    ru_holidays = holidays.RU(years=[current_year - 2, current_year - 1, current_year,
                                     current_year + 1, current_year + 2, current_year + 3])
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


# 1. Извлечение вехи/даты старта из "START PROJECT TOGF-ENG-007-02"
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


# 1.1. Извлечение списка вех из "START PROJECT TOGF-ENG-007-02"
#      Колонка B (индекс 1) — названия, строки 30–35
#      Колонка C (индекс 2) — даты,     строки 30–35
def extract_milestones_from_start_sheet(xls):
    sheet_name = "START PROJECT TOGF-ENG-007-02"
    milestones = []
    if sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        # строки 30–35 в 1-индексации => индексы 29–34 в pandas
        for r in range(29, 35):
            if r >= df.shape[0]:
                break
            name_val = df.iloc[r, 1] if df.shape[1] > 1 else None  # колонка B
            date_val = df.iloc[r, 2] if df.shape[1] > 2 else None  # колонка C
            if pd.notna(name_val) and str(name_val).strip():
                dt = pd.to_datetime(date_val, errors='coerce')
                milestones.append({
                    'Name': str(name_val).strip(),
                    'Date': dt.date() if pd.notna(dt) else None,
                    'Source': sheet_name
                })
    return milestones


# 1.2. Извлечение вех-маркеров из "PMSPR TOGF-ENG-008-06 Phase 3"
#      Колонка BK (индекс 63) — строка 9 (индекс 8) — название
#      Колонка BK (индекс 63) — строка 10 (индекс 9) — дата
def extract_phase3_milestones(xls):
    sheet_name = "PMSPR TOGF-ENG-008-06 Phase 3"
    milestones = []
    if sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        col_idx = 63  # BK
        if df.shape[1] > col_idx:
            name_val = df.iloc[8, col_idx] if df.shape[0] > 8 else None   # строка 9
            date_val = df.iloc[9, col_idx] if df.shape[0] > 9 else None   # строка 10
            if pd.notna(name_val) and str(name_val).strip():
                dt = pd.to_datetime(date_val, errors='coerce')
                milestones.append({
                    'Name': str(name_val).strip(),
                    'Date': dt.date() if pd.notna(dt) else None,
                    'Source': sheet_name
                })
    return milestones


uploaded_file = st.file_uploader("Загрузите файл (PLANT_MASTER_SCHEDULE P25077.xlsx)", type=["xlsx"])

# 2. Последовательность вкладок
target_phases = [
    "PMSPR TOGF-ENG-008-06 Phase2",
    "PMSPR TOGF-ENG-008-06 Phase 3",
    "PMSPR TOGF-ENG-008-06 Phase4;5"
]

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        start_date = extract_start_milestone(xls)

        if start_date:
            st.success(f"📅 Дата вехи извлечена из листа 'START PROJECT TOGF-ENG-007-02': **{start_date.strftime('%d.%m.%Y')}**")

            if st.button("🚀 Сформировать автоматический SMS-график"):
                with st.spinner("Формирование графика и построение диаграммы Ганта..."):
                    tasks = []

                    # Извлечение DESCRIPTION из указанных вкладок по порядку
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
                        df_tasks = pd.DataFrame(tasks)
                        holiday_dates = get_rf_holidays()

                        current_date = start_date
                        schedule = []
                        gantt_data = []

                        # Расчет дат с учетом рабочих дней РФ
                        for idx, row in df_tasks.iterrows():
                            current_date = get_next_business_day(current_date, holiday_dates)

                            t_start = current_date
                            t_end = t_start + datetime.timedelta(days=1)  # для отображения полосы в Ганте

                            task_name = (f"{idx + 1}. {row['DESCRIPTION'][:60]}..."
                                         if len(row['DESCRIPTION']) > 60
                                         else f"{idx + 1}. {row['DESCRIPTION']}")

                            schedule.append({
                                '№': idx + 1,
                                'Фаза проекта': row['Phase'],
                                'DESCRIPTION': row['DESCRIPTION'],
                                'Дата начала (расчет)': t_start.strftime('%d.%m.%Y'),
                                'Дата окончания (расчет)': t_start.strftime('%d.%m.%Y')
                            })

                            gantt_data.append({
                                'Task': task_name,
                                'Start': t_start,
                                'Finish': t_end,
                                'Phase': row['Phase'],
                                'Full_Description': row['DESCRIPTION']
                            })

                            current_date = get_next_business_day(t_start + datetime.timedelta(days=1), holiday_dates)

                        res_df = pd.DataFrame(schedule)
                        gantt_df = pd.DataFrame(gantt_data)

                        # ---- Сбор вех ----
                        milestones = []
                        milestones += extract_milestones_from_start_sheet(xls)
                        milestones += extract_phase3_milestones(xls)

                        # убираем вехи без даты
                        milestones = [m for m in milestones if m['Date'] is not None]

                        # ---- Диаграмма Ганта ----
                        st.subheader("📊 Диаграмма Ганта проекта")

                        fig = px.timeline(
                            gantt_df,
                            x_start="Start",
                            x_end="Finish",
                            y="Task",
                            color="Phase",
                            hover_data=["Full_Description"],
                            title="Календарный SMS-график запуска в серийное производство"
                        )
                        fig.update_yaxes(autorange="reversed")

                        # ---- Наложение вех как вертикальных линий/маркеров ----
                        if milestones:
                            # Определяем границы по X, чтобы рисовать вертикальные линии на всю высоту
                            x_min = gantt_df['Start'].min()
                            x_max = gantt_df['Finish'].max()

                            for i, m in enumerate(milestones):
                                m_date = pd.to_datetime(m['Date'])

                                # Вертикальная линия вехи
                                fig.add_vline(
                                    x=m_date,
                                    line_width=2,
                                    line_dash="dash",
                                    line_color="crimson"
                                )

                                # Подпись вехи (сверху диаграммы)
                                fig.add_annotation(
                                    x=m_date,
                                    y=1.02,
                                    yref="paper",
                                    text=f"🚩 {m['Name']}<br>{m_date.strftime('%d.%m.%Y')}",
                                    showarrow=False,
                                    font=dict(size=10, color="crimson"),
                                    bgcolor="rgba(255,255,255,0.85)",
                                    bordercolor="crimson",
                                    borderwidth=1,
                                    align="center"
                                )

                                # Ромб-маркер на самой линии (для наглядности)
                                fig.add_trace(go.Scatter(
                                    x=[m_date],
                                    y=[gantt_df['Task'].iloc[0]],
                                    mode="markers",
                                    marker=dict(symbol="diamond", size=14,
                                                color="crimson",
                                                line=dict(color="white", width=1)),
                                    name=f"🚩 {m['Name']}",
                                    hovertemplate=(f"<b>{m['Name']}</b><br>"
                                                   f"Дата: {m_date.strftime('%d.%m.%Y')}<br>"
                                                   f"Источник: {m['Source']}<extra></extra>"),
                                    showlegend=True
                                ))

                        fig.update_layout(
                            height=max(500, len(gantt_df) * 25),
                            xaxis_title="Дата",
                            yaxis_title="Задачи (DESCRIPTION)",
                            legend_title="Фаза проекта / Вехи",
                            margin=dict(t=120)  # место под подписи вех сверху
                        )

                        st.plotly_chart(fig, use_container_width=True)

                        # ---- Таблица вех ----
                        if milestones:
                            st.subheader("🚩 Вехи проекта")
                            ms_df = pd.DataFrame([{
                                'Веха': m['Name'],
                                'Дата': m['Date'].strftime('%d.%m.%Y'),
                                'Источник': m['Source']
                            } for m in milestones])
                            st.dataframe(ms_df, use_container_width=True)

                        # ---- Таблица календарного плана ----
                        st.subheader("📋 Таблица календарного плана")
                        st.dataframe(res_df, use_container_width=True)

                        # ---- Экспорт в Excel ----
                        excel_out = "SMS_Schedule_Result.xlsx"
                        with pd.ExcelWriter(excel_out, engine='openpyxl') as writer:
                            res_df.to_excel(writer, index=False, sheet_name='SMS Schedule')
                            if milestones:
                                ms_export = pd.DataFrame([{
                                    'Веха': m['Name'],
                                    'Дата': m['Date'].strftime('%d.%m.%Y'),
                                    'Источник': m['Source']
                                } for m in milestones])
                                ms_export.to_excel(writer, index=False, sheet_name='Milestones')

                        with open(excel_out, "rb") as f:
                            st.download_button("📥 Скачать итоговый Excel", f,
                                               file_name="SMS_Schedule_Result.xlsx")
                    else:
                        st.error("❌ Не удалось извлечь задачи из столбцов DESCRIPTION.")
        else:
            st.error("❌ Не удалось найти дату старта на вкладке 'START PROJECT TOGF-ENG-007-02'.")

    except Exception as e:
        st.error(f"Ошибка при обработке файла: {e}")
else:
    st.info("ℹ️ Для запуска расчета загрузите Excel-файл.")
