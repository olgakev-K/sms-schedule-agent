import streamlit as st
import pandas as pd
import openpyxl
from openpyxl.styles.colors import Color
import plotly.express as px
import plotly.graph_objects as go
from datetime import datetime, timedelta
from workalendar.europe import Russia
import io

# --- НАСТРОЙКИ СТРАНИЦЫ ---
st.set_page_config(page_title="Генератор SMS Графика", layout="wide")
st.title("🏭 Генератор SMS Графика Запуска (Plant Master Schedule)")
st.markdown("Загрузите файл шаблона **PLANT_MASTER_SCHEDULE P25077**, и ИИ-агент автоматически сформирует календарный план с учетом производственного календаря РФ.")

# --- ВСПОМОГАТЕЛЬНЫЕ ФУНКЦИИ ---
def get_working_days_duration(weeks: int, start_date: datetime, calendar: Russia) -> datetime:
    """Рассчитывает дату окончания, добавляя рабочие недели и пропуская выходные/праздники РФ."""
    working_days_to_add = weeks * 5  # 5 рабочих дней в неделе
    return calendar.add_working_days(start_date, working_days_to_add)

def count_colored_cells_in_row(row, start_col_index: int = 38): # 38 = колонка AM (A=0, B=1... AM=38)
    """Считает количество ячеек с цветной заливкой в строке, начиная с указанной колонки."""
    count = 0
    for cell in row[start_col_index:]:
        # Проверяем, есть ли сплошная заливка и цвет не является прозрачным/белым по умолчанию
        if cell.fill.patternType == 'solid' and cell.fill.start_color.index != '00000000' and cell.fill.start_color.index != 'FFFFFFFF':
            count += 1
    return count

def process_excel(uploaded_file):
    """Основная логика обработки Excel файла."""
    cal = Russia()
    
    # Загружаем книгу через openpyxl для сохранения информации о цветах
    wb = openpyxl.load_workbook(uploaded_file, data_only=False)
    
    # 1. ИЗВЛЕЧЕНИЕ ВЕХ (Milestones)
    milestones = []
    try:
        ws_start = wb["START PROJECT TOGF-ENG-007-02"]
        for row_idx in range(30, 36): # Строки 30-35 (1-indexed)
            name_cell = ws_start.cell(row=row_idx, column=2) # Колонка B
            date_cell = ws_start.cell(row=row_idx, column=3) # Колонка C
            
            if name_cell.value and date_cell.value:
                # Преобразуем дату Excel в datetime
                m_date = date_cell.value
                if isinstance(m_date, datetime):
                    milestones.append({"Name": str(name_cell.value).strip(), "Date": m_date})
    except KeyError:
        st.error("Не найдена вкладка 'START PROJECT TOGF-ENG-007-02'")
        return None

    # Сортируем вехи по дате для определения базовой точки отсчета
    milestones.sort(key=lambda x: x["Date"])
    base_start_date = milestones[0]["Date"] if milestones else datetime.now()

    # 2. ИЗВЛЕЧЕНИЕ ЗАДАЧ (Tasks)
    tasks_data = []
    phase_sheets = ["PMSPR TOGF-ENG-008-06 Phase2", "PMSPR TOGF-ENG-008-06 Phase 3", "PMSPR TOGF-ENG-008-06 Phase4;5"]
    
    for sheet_name in phase_sheets:
        try:
            ws = wb[sheet_name]
            # Проходим по строкам (предполагаем, что данные начинаются примерно с 10-й строки, настройте при необходимости)
            for row_idx in range(10, 100): 
                desc_cell = ws.cell(row=row_idx, column=4) # Предположим, DESCRIPTION в колонке D (индекс 4). Уточните под ваш файл!
                
                if desc_cell.value and isinstance(desc_cell.value, str) and len(desc_cell.value.strip()) > 3:
                    duration_weeks = count_colored_cells_in_row(ws[row_idx], start_col_index=38) # Колонка AM
                    
                    if duration_weeks > 0:
                        # Логика привязки: берем ближайшую прошедшую или текущую веху как старт
                        # Для упрощения: старт = base_start_date + смещение (в реальном проекте тут нужна логика зависимостей)
                        start_date = base_start_date + timedelta(days=len(tasks_data) * 2) # Заглушка для последовательности
                        
                        end_date = get_working_days_duration(duration_weeks, start_date, cal)
                        
                        tasks_data.append({
                            "Фаза": sheet_name,
                            "Описание (DESCRIPTION)": desc_cell.value.strip(),
                            "Duration (weeks)": duration_weeks,
                            "Start Date": start_date,
                            "End Date": end_date
                        })
        except KeyError:
            st.warning(f"Вкладка '{sheet_name}' не найдена, пропускаем.")

    return pd.DataFrame(tasks_data), milestones

# --- ИНТЕРФЕЙС STREAMLIT ---
uploaded_file = st.file_uploader("Загрузите файл PLANT_MASTER_SCHEDULE P25077.xlsx", type=["xlsx", "xls"])

if uploaded_file is not None:
    with st.spinner("🤖 ИИ-агент анализирует файл, считает цветные ячейки и строит календарь РФ..."):
        result = process_excel(uploaded_file)
        
        if result:
            df_tasks, milestones = result
            
            st.success("✅ Обработка завершена успешно!")
            
            # Колонки для отображения
            col1, col2 = st.columns([1, 2])
            
            with col1:
                st.subheader("📋 Таблица SMS")
                # Форматируем даты для красивого вывода
                df_display = df_tasks.copy()
                df_display["Start Date"] = df_display["Start Date"].dt.strftime("%d.%m.%Y")
                df_display["End Date"] = df_display["End Date"].dt.strftime("%d.%m.%Y")
                st.dataframe(df_display, use_container_width=True, hide_index=True)
                
                st.subheader("🚩 Вехи проекта")
                df_milestones = pd.DataFrame(milestones)
                df_milestones["Date"] = df_milestones["Date"].dt.strftime("%d.%m.%Y")
                st.dataframe(df_milestones, use_container_width=True, hide_index=True)

            with col2:
                st.subheader("📊 Диаграмма Ганта с маркерами вех")
                
                # Подготовка данных для Plotly
                fig = px.timeline(
                    df_tasks,
                    x_start="Start Date",
                    x_end="End Date",
                    y="Описание (DESCRIPTION)",
                    color="Фаза",
                    title="График запуска в серийное производство",
                    hover_data=["Duration (weeks)"]
                )
                
                # Инвертируем ось Y, чтобы первая задача была сверху
                fig.update_yaxes(autorange="reversed")
                
                # ДОБАВЛЕНИЕ ВЕХ НА ГРАФИК (Вертикальные линии с маркерами)
                for m in milestones:
                    fig.add_vline(
                        x=m["Date"],
                        line_dash="dash",
                        line_color="red",
                        line_width=2,
                        annotation_text=f"🚩 {m['Name']}",
                        annotation_position="top right"
                    )
                    # Добавляем маркер-ромб (◆) как в колонке BK шаблона
                    fig.add_trace(go.Scatter(
                        x=[m["Date"]],
                        y=[df_tasks["Описание (DESCRIPTION)"].iloc[0]], # Привязываем к первой строке для высоты
                        mode="markers",
                        marker=dict(symbol="diamond", size=12, color="red"),
                        showlegend=False,
                        hoverinfo="text",
                        text=m["Name"]
                    ))

                fig.update_layout(
                    xaxis_title="Календарный план (с учетом выходных и праздников РФ)",
                    yaxis_title="Задачи",
                    height=max(400, len(df_tasks) * 25), # Динамическая высота
                    xaxis=dict(tickformat="%d.%m.%Y")
                )
                
                st.plotly_chart(fig, use_container_width=True)
                
                # Кнопка скачивания результата
                csv = df_display.to_csv(index=False, sep=";").encode('utf-8-sig')
                st.download_button(
                    label="📥 Скачать SMS в CSV",
                    data=csv,
                    file_name='SMS_Schedule_RF_Calendar.csv',
                    mime='text/csv',
                )
