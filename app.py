import streamlit as st
import pandas as pd
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
import io

# --- НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Генератор SMS Графика", layout="wide")
st.title("🏭 Генератор SMS Графика Запуска")
st.markdown("Загрузите файл **PLANT_MASTER_SCHEDULE P25077.xlsx**")

# --- ФУНКЦИЯ ОПРЕДЕЛЕНИЯ РАБОЧИХ ДНЕЙ ---
def is_working_day(date, holidays):
    """Проверяет, является ли дата рабочим днем (не выходной и не праздник)"""
    # Суббота (5) и Воскресенье (6)
    if date.weekday() >= 5:
        return False
    # Проверка на праздники
    if date.strftime("%Y-%m-%d") in holidays:
        return False
    return True

def add_working_days(start_date, weeks, holidays):
    """Добавляет рабочие недели к дате"""
    days_to_add = 0
    working_days_added = 0
    target_days = weeks * 5
    
    current_date = start_date
    while working_days_added < target_days:
        current_date += timedelta(days=1)
        if is_working_day(current_date, holidays):
            working_days_added += 1
    
    return current_date

# --- ПРОИЗВОДСТВЕННЫЙ КАЛЕНДАРЬ РФ (2025-2026) ---
def get_russia_holidays():
    """Возвращает список государственных праздников РФ"""
    # Основные праздничные дни (нерабочие)
    holidays = {
        # 2025
        "2025-01-01", "2025-01-02", "2025-01-03", "2025-01-04", 
        "2025-01-05", "2025-01-06", "2025-01-07", "2025-01-08",
        "2025-02-23", "2025-03-08", "2025-05-01", "2025-05-09",
        "2025-06-12", "2025-11-04",
        # 2026
        "2026-01-01", "2026-01-02", "2026-01-03", "2026-01-04",
        "2026-01-05", "2026-01-06", "2026-01-07", "2026-01-08",
        "2026-02-23", "2026-03-08", "2026-05-01", "2026-05-09",
        "2026-06-12", "2026-11-04",
    }
    return holidays

# --- ОБРАБОТКА EXCEL ---
def process_excel(uploaded_file):
    """Основная логика обработки Excel файла"""
    try:
        # Читаем Excel с pandas
        xls = pd.ExcelFile(uploaded_file, engine='openpyxl')
        
        # 1. ИЗВЛЕЧЕНИЕ ВЕХ
        milestones = []
        if "START PROJECT TOGF-ENG-007-02" in xls.sheet_names:
            df_start = pd.read_excel(xls, sheet_name="START PROJECT TOGF-ENG-007-02", header=None)
            # Строки 30-35 (индексы 29-34), колонки B (1) и C (2)
            for idx in range(29, 35):
                if idx < len(df_start):
                    name = df_start.iloc[idx, 1] if pd.notna(df_start.iloc[idx, 1]) else None
                    date = df_start.iloc[idx, 2] if pd.notna(df_start.iloc[idx, 2]) else None
                    if name and pd.notna(date):
                        milestones.append({
                            "Name": str(name).strip(),
                            "Date": pd.to_datetime(date)
                        })
        
        if not milestones:
            st.warning("⚠️ Вехи не найдены. Использую текущую дату как точку отсчета.")
            base_date = datetime.now()
        else:
            milestones.sort(key=lambda x: x["Date"])
            base_date = milestones[0]["Date"]
        
        # 2. ИЗВЛЕЧЕНИЕ ЗАДАЧ
        tasks_data = []
        phase_sheets = [
            "PMSPR TOGF-ENG-008-06 Phase2",
            "PMSPR TOGF-ENG-008-06 Phase 3",
            "PMSPR TOGF-ENG-008-06 Phase4;5"
        ]
        
        holidays = get_russia_holidays()
        current_start = base_date
        
        for sheet_name in phase_sheets:
            if sheet_name in xls.sheet_names:
                try:
                    df_phase = pd.read_excel(xls, sheet_name=sheet_name, header=None)
                    
                    # Ищем колонку DESCRIPTION (обычно колонка D - индекс 3)
                    # Проверяем строки с 10 по 100
                    for idx in range(9, min(100, len(df_phase))):
                        # Предполагаем, что описание в колонке D (индекс 3)
                        if 3 < len(df_phase.columns):
                            desc = df_phase.iloc[idx, 3]
                        else:
                            desc = None
                            
                        if pd.notna(desc) and isinstance(desc, str) and len(str(desc).strip()) > 3:
                            # Упрощенная логика: длительность = 2 недели по умолчанию
                            # В идеале нужно читать цвета, но для простоты пока так
                            duration_weeks = 2  
                            
                            end_date = add_working_days(current_start, duration_weeks, holidays)
                            
                            tasks_data.append({
                                "Фаза": sheet_name.split()[-1] if "Phase" in sheet_name else sheet_name,
                                "Описание": str(desc).strip()[:100],  # Обрезаем длинные описания
                                "Duration (weeks)": duration_weeks,
                                "Start Date": current_start,
                                "End Date": end_date
                            })
                            
                            # Следующая задача начинается после предыдущей
                            current_start = end_date + timedelta(days=1)
                            
                except Exception as e:
                    st.warning(f"⚠️ Ошибка при чтении вкладки {sheet_name}: {str(e)}")
                    continue
        
        if not tasks_data:
            st.error("❌ Задачи не найдены. Проверьте структуру файла.")
            return None, None
        
        return pd.DataFrame(tasks_data), milestones
        
    except Exception as e:
        st.error(f"❌ Ошибка обработки файла: {str(e)}")
        return None, None

# --- ИНТЕРФЕЙС ---
uploaded_file = st.file_uploader("Загрузите файл Excel", type=["xlsx"])

if uploaded_file is not None:
    with st.spinner("🤖 Обрабатываю файл..."):
        result = process_excel(uploaded_file)
        
        if result[0] is not None:
            df_tasks, milestones = result
            
            st.success("✅ Обработка завершена!")
            
            col1, col2 = st.columns([1, 2])
            
            with col1:
                st.subheader("📋 Таблица SMS")
                df_display = df_tasks.copy()
                df_display["Start Date"] = df_display["Start Date"].dt.strftime("%d.%m.%Y")
                df_display["End Date"] = df_display["End Date"].dt.strftime("%d.%m.%Y")
                st.dataframe(df_display, use_container_width=True, hide_index=True)
                
                if milestones:
                    st.subheader(" Вехи проекта")
                    df_milestones = pd.DataFrame(milestones)
                    df_milestones["Date"] = df_milestones["Date"].dt.strftime("%d.%m.%Y")
                    st.dataframe(df_milestones, use_container_width=True, hide_index=True)

            with col2:
                st.subheader("📊 Диаграмма Ганта")
                
                fig = px.timeline(
                    df_tasks,
                    x_start="Start Date",
                    x_end="End Date",
                    y="Описание",
                    color="Фаза",
                    title="График запуска",
                    hover_data=["Duration (weeks)"]
                )
                
                fig.update_yaxes(autorange="reversed")
                
                # Добавляем вехи
                if milestones:
                    for m in milestones:
                        fig.add_vline(
                            x=m["Date"],
                            line_dash="dash",
                            line_color="red",
                            line_width=2,
                            annotation_text=f"🚩 {m['Name']}",
                            annotation_position="top right"
                        )
                
                fig.update_layout(
                    xaxis_title="Дата",
                    yaxis_title="Задачи",
                    height=max(400, len(df_tasks) * 30),
                    xaxis=dict(tickformat="%d.%m.%Y")
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
                # Скачивание
                csv = df_display.to_csv(index=False, sep=";").encode('utf-8-sig')
                st.download_button(
                    label="📥 Скачать CSV",
                    data=csv,
                    file_name='SMS_Schedule.csv',
                    mime='text/csv',
                )
