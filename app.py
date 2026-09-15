import streamlit as st
import pandas as pd
import datetime
import requests
from bs4 import BeautifulSoup
import holidays
import plotly.express as px
import plotly.graph_objects as go
from openpyxl.styles import Font, Alignment
from openpyxl.utils import get_column_letter

st.set_page_config(page_title="SMS Schedule Agent", layout="wide")

st.title("🤖 ИИ-Агент: Генератор SMS-графика проекта")
st.write("Автоматический расчет календарного плана с визуализацией в виде Диаграммы Ганта")


# ---------- Праздники РФ ----------
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


def add_business_days(start_date, n_days, holiday_dates):
    """Возвращает дату, отстоящую от start_date на n рабочих дней (n_days >= 0).
    n_days=0 -> первый рабочий день, начиная с start_date."""
    cur = start_date
    added = 0
    while True:
        cur = get_next_business_day(cur, holiday_dates)
        if added == n_days:
            return cur
        cur += datetime.timedelta(days=1)
        added += 1


# ---------- Извлечение даты старта ----------
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


# ---------- Вехи из START-листа ----------
def extract_milestones_from_start_sheet(xls):
    sheet_name = "START PROJECT TOGF-ENG-007-02"
    milestones = []
    if sheet_name in xls.sheet_names:
        df = pd.read_excel(xls, sheet_name=sheet_name, header=None)
        for r in range(29, 35):  # строки 30–35
            if r >= df.shape[0]:
                break
            name_val = df.iloc[r, 1] if df.shape[1] > 1 else None  # B
            date_val = df.iloc[r, 2] if df.shape[1] > 2 else None  # C
            if pd.notna(name_val) and str(name_val).strip():
                dt = pd.to_datetime(date_val, errors='coerce')
                milestones.append({
                    'Name': str(name_val).strip(),
                    'Date': dt.date() if pd.notna(dt) else None,
                    'Source': sheet_name
                })
    return milestones


# ---------- Вехи из Phase 3 (BK9 / BK10) ----------
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

target_phases = [
    "PMSPR TOGF-ENG-008-06 Phase2",
    "PMSPR TOGF-ENG-008-06 Phase 3",
    "PMSPR TOGF-ENG-008-06 Phase4;5"
]

# Длительность по умолчанию для одной задачи (в рабочих днях)
TASK_DURATION_WORKDAYS = 1

if uploaded_file:
    try:
        xls = pd.ExcelFile(uploaded_file)
        start_date = extract_start_milestone(xls)

        if start_date:
            st.success(f"📅 Дата вехи извлечена из листа 'START PROJECT TOGF-ENG-007-02': "
                       f"**{start_date.strftime('%d.%m.%Y')}**")

            if st.button("🚀 Сформировать автоматический SMS-график"):
                with st.spinner("Формирование графика и построение диаграммы Ганта..."):
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
                        df_tasks = pd.DataFrame(tasks)
                        holiday_dates = get_rf_holidays()

                        schedule = []
                        gantt_data = []

                        # --- корректный расчёт дат ---
                        # Каждая задача начинается с первого рабочего дня после предыдущей
                        # и длится TASK_DURATION_WORKDAYS рабочих дней.
                        cursor = start_date  # стартовая точка (нерабочий день допустим — сдвинется)

                        for idx, row in df_tasks.iterrows():
                            # начало — следующий рабочий день от cursor
                            t_start = get_next_business_day(cursor, holiday_dates)

                            # окончание — последний рабочий день выполнения
                            t_end = add_business_days(t_start, TASK_DURATION_WORKDAYS - 1, holiday_dates)

                            # следующий поиск начинается со дня после окончания
                            cursor = t_end + datetime.timedelta(days=1)

                            task_name = (f"{idx + 1}. {row['DESCRIPTION'][:60]}..."
                                         if len(row['DESCRIPTION']) > 60
                                         else f"{idx + 1}. {row['DESCRIPTION']}")

                            schedule.append({
                                '№': idx + 1,
                                'Фаза проекта': row['Phase'],
                                'DESCRIPTION': row['DESCRIPTION'],
                                'Дата начала': t_start.strftime('%d.%m.%Y'),
                                'Дата окончания': t_end.strftime('%d.%m.%Y')
                            })

                            # Для Plotly x_end — эксклюзивная граница, поэтому +1 день
                            gantt_data.append({
                                'Task': task_name,
                                'Start': pd.to_datetime(t_start),
                                'Finish': pd.to_datetime(t_end) + pd.Timedelta(days=1),
                                'Phase': row['Phase'],
                                'Full_Description': row['DESCRIPTION']
                            })

                        res_df = pd.DataFrame(schedule)
                        gantt_df = pd.DataFrame(gantt_data)

                        # ---- Вехи ----
                        milestones = []
                        milestones += extract_milestones_from_start_sheet(xls)
                        milestones += extract_phase3_milestones(xls)
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

                        # ---- Отрисовка вех ----
                        if milestones:
                            # Сдвигаем подписи по вертикали, если даты совпадают
                            used_dates = {}
                            for m in milestones:
                                m_date = pd.to_datetime(m['Date'])
                                offset = used_dates.get(m_date, 0)
                                used_dates[m_date] = offset + 1

                                fig.add_vline(
                                    x=m_date,
                                    line_width=2,
                                    line_dash="dash",
                                    line_color="crimson"
                                )
                                fig.add_annotation(
                                    x=m_date,
                                    y=1.02 + offset * 0.06,
                                    yref="paper",
                                    text=f"🚩 {m['Name']}<br>{m_date.strftime('%d.%m.%Y')}",
                                    showarrow=False,
                                    font=dict(size=10, color="crimson"),
                                    bgcolor="rgba(255,255,255,0.85)",
                                    bordercolor="crimson",
                                    borderwidth=1,
                                    align="center"
                                )

                        fig.update_layout(
                            height=max(500, len(gantt_df) * 25),
                            xaxis_title="Дата",
                            yaxis_title="Задачи (DESCRIPTION)",
                            legend_title="Фаза проекта",
                            margin=dict(t=140)
                        )

                        st.plotly_chart(fig, use_container_width=True)

                        # ---- Таблица вех (на экране) ----
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

                        # ---- Excel: один лист, первая строка жирным ----
                        excel_out = "SMS_Schedule_Result.xlsx"

                        # Объединяем задачи и вехи в один датафрейм
                        combined_rows = []
                        for _, r in res_df.iterrows():
                            combined_rows.append({
                                '№': r['№'],
                                'Фаза проекта': r['Фаза проекта'],
                                'DESCRIPTION / Веха': r['DESCRIPTION'],
                                'Дата начала': r['Дата начала'],
                                'Дата окончания': r['Дата окончания'],
                                'Тип': 'Задача'
                            })
                        for m in milestones:
                            combined_rows.append({
                                '№': '',
                                'Фаза проекта': 'Веха',
                                'DESCRIPTION / Веха': f"🚩 {m['Name']}",
                                'Дата начала': m['Date'].strftime('%d.%m.%Y'),
                                'Дата окончания': m['Date'].strftime('%d.%m.%Y'),
                                'Тип': 'Веха'
                            })

                        export_df = pd.DataFrame(combined_rows)

                        with pd.ExcelWriter(excel_out, engine='openpyxl') as writer:
                            export_df.to_excel(writer, index=False, sheet_name='SMS Schedule')
                            ws = writer.sheets['SMS Schedule']

                            # Первая строка — жирным + выравнивание по центру
                            for cell in ws[1]:
                                cell.font = Font(bold=True)
                                cell.alignment = Alignment(horizontal='center', vertical='center')

                            # Автоширина колонок
                            for col_idx, column in enumerate(export_df.columns, start=1):
                                max_len = max(
                                    [len(str(column))] +
                                    [len(str(v)) for v in export_df[column].astype(str).tolist()]
                                )
                                ws.column_dimensions[get_column_letter(col_idx)].width = min(max_len + 3, 70)

                            # Закрепить шапку
                            ws.freeze_panes = "A2"

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
