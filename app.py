import plotly.express as px
import pandas as pd

# Данные из Excel
df = pd.read_excel('project.xlsx')

# Расчёт дат
df['Start'] = pd.to_datetime(df['Start_Date'])
df['Finish'] = df['Start'] + pd.to_timedelta(df['Duration_Days'], unit='D')

# Диаграмма Ганта
fig = px.timeline(df, x_start="Start", x_end="Finish", y="Task")
fig.update_yaxes(autorange="reversed")
fig.update_layout(xaxis_title="Недели проекта")

# Группировка по неделям
fig.update_xaxes(
    dtick=7*24*60*60*1000,  # 1 неделя в миллисекундах
    tickformat="Неделя %W"
)

fig.show()
