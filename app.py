# app.py — SMS-график запуска в серийное производство (P25077)
# Запуск: streamlit run app.py
# Установка: pip install streamlit openpyxl

import datetime
import io

import streamlit as st
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

st.set_page_config(page_title='SMS-график P25077', layout='centered')
st.title('📋 SMS-график запуска в серийное производство')
st.caption('Источник: PLANT_MASTER_SCHEDULE → вкладки Phase2 → Phase3 → Phase4;5. '
           'Длительность = закрашенные недели Ганта (item end − item start + 1). '
           'Календарь: только рабочие дни РФ (без выходных и госпраздников).')

# ---------------- 1. Загрузка файла ----------------
uploaded = st.file_uploader(
    'Загрузите файл PLANT_MASTER_SCHEDULE P25077.xlsx '
    '(⚠️ предварительно откройте его в Excel и сохраните — иначе кэш формул пуст)',
    type=['xlsx'])
if uploaded is None:
    st.info('Ожидаю файл…')
    st.stop()

# ---------------- 2. Производственный календарь РФ ----------------
HOL = set()
for y in range(2025, 2033):
    HOL |= {datetime.date(y, 1, d) for d in range(1, 9)}
    HOL |= {datetime.date(y, 2, 23), datetime.date(y, 3, 8),
            datetime.date(y, 5, 1), datetime.date(y, 5, 2),
            datetime.date(y, 5, 9), datetime.date(y, 6, 12),
            datetime.date(y, 11, 4), datetime.date(y, 12, 31)}
HOL |= {datetime.date(2025, 3, 9), datetime.date(2026, 3, 9)}  # 8 марта в выходной
# TODO: сверить официальные переносы выходных на годы проекта (consultant.ru)

def is_work(d):
    return d.weekday() < 5 and d not in HOL

def next_work(d):
    while not is_work(d):
        d += datetime.timedelta(days=1)
    return d

def add_workdays(start, n):
    """Дата n-го рабочего дня, считая start первым."""
    d, c = start, 1
    while c < n:
        d += datetime.timedelta(days=1)
        if is_work(d):
            c += 1
    return d

def to_date(v):
    if isinstance(v, datetime.datetime):
        return v.date()
    if isinstance(v, (int, float)):
        return datetime.date(1899, 12, 30) + datetime.timedelta(days=int(v))
    return None

# ---------------- 3. Чтение исходника ----------------
with st.spinner('Читаю вкладки Phase 2 → Phase 3 → Phase 4;5…'):
    wbv = openpyxl.load_workbook(uploaded, data_only=True)

CFG = [('PMSPR TOGF-ENG-008-06 Phase2', 36, 'Фаза 2 — Разработка продукта и процесса'),
       ('PMSPR TOGF-ENG-008-06 Phase 3', 39, 'Фаза 3 — Индустриализация'),
       ('PMSPR TOGF-ENG-008-06 Phase4;5', 39, 'Фаза 4/5 — Наращивание / Серийная жизнь')]

items = []          # [№, Description, Responsible, Phase, DurationWeeks]
for name, resp_col, phase in CFG:
    if name not in wbv.sheetnames:
        st.error(f'Вкладка не найдена: {name}')
        st.stop()
    ws = wbv[name]
    for r in range(13, ws.max_row + 1):
        desc = ws.cell(r, 25).value                       # колонка Y — DESCRIPTION
        if not (isinstance(desc, str) and desc.strip()):
            continue
        c_start = ws.cell(r, 3).value                     # служебная: item start (неделя)
        c_end = ws.cell(r, 12).value                      # служебная: item end (неделя)
        dur = None
        if isinstance(c_start, (int, float)) and isinstance(c_end, (int, float)) and c_end >= c_start:
            dur = int(round(c_end - c_start + 1))         # = кол-во закрашенных ячеек Ганта
        items.append([str(ws.cell(r, 24).value or '').strip(),  # № из колонки X
                      desc.strip(), ws.cell(r, resp_col).value or '', phase, dur])

missing = sum(1 for i in items if i[4] is None)
if missing:
    st.error(f'⚠️ У {missing} из {len(items)} строк пуст кэш формул. '
             'Откройте файл в Excel, нажмите Ctrl+S и загрузите снова.')
    st.stop()

# ---------------- 4. Вехи и точка отсчёта ----------------
ms_ws = wbv['START PROJECT TOGF-ENG-007-02']
milestones, ms_rows = {}, []
for r in range(1, ms_ws.max_row + 1):
    b = ms_ws.cell(r, 2).value
    if isinstance(b, str) and b.strip() in ('VC', 'PT1', 'PT2', 'PPC', 'PSW', 'SOP'):
        milestones[b.strip()] = to_date(ms_ws.cell(r, 3).value)
        ms_rows.append((r, b.strip(), milestones[b.strip()]))

phase2_start = to_date(wbv['PMSPR TOGF-ENG-008-06 Phase2'].cell(13, 1).value) \
               or datetime.date(2025, 6, 11)

with st.expander('Проверка вех (вкладка START PROJECT, колонки B/C)'):
    st.write('Строки:', ms_rows)
    st.warning('⚠️ Даты вех 2028–2031 выглядят как неотредактированный лист начала проекта '
               '(сам график живёт в 2025–2027). Проверьте у заказчика.')
    st.write('Старт Фазы 2 (точка отсчёта):', phase2_start)

# ---------------- 5. Расчёт календаря (цепочка, рабочие дни РФ) ----------------
cur = next_work(phase2_start)
for it in items:
    wd = max(1, int(round(it[4] * 5)))                    # недели → рабочие дни (5-дневка)
    start = cur
    finish = add_workdays(start, wd)
    it += [start, finish]
    cur = next_work(finish + datetime.timedelta(days=1))

# ---------------- 6. Сборка SMS: таблица + Ганта на одном листе ----------------
out = openpyxl.Workbook()
ws = out.active
ws.title = 'SMS P25077'

HDR = ['№', 'Description', 'Phase', 'Duration, weeks', 'Start', 'Finish']
head_fill = PatternFill('solid', fgColor='4472C4')
for c, h in enumerate(HDR, 1):
    cell = ws.cell(3, c, h)
    cell.font = Font(bold=True, color='FFFFFF')
    cell.fill = head_fill
    cell.alignment = Alignment(horizontal='center', wrap_text=True)

bar_fill = PatternFill('solid', fgColor='2E75B6')
mile_fill = PatternFill('solid', fgColor='FFC000')
thin = Border(*[Side(style='thin', color='D9D9D9')] * 4)

w0 = items[0][5] - datetime.timedelta(days=items[0][5].weekday())   # понедельник 1-й недели
weeks = ((items[-1][6] - w0).days // 7) + 2
for wcol in range(weeks):
    mon = w0 + datetime.timedelta(weeks=wcol)
    cell = ws.cell(2, 7 + wcol, f'CW{mon.isocalendar().week:02d}/{mon.year}')
    cell.font = Font(size=7)
    cell.alignment = Alignment(horizontal='center', textRotation=90)

for i, it in enumerate(items):
    r = 4 + i
    for c, v in enumerate(it[:6], 1):
        cell = ws.cell(r, c, v)
        cell.border = thin
        if c in (5, 6):
            cell.number_format = 'DD.MM.YYYY'
    ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical='center')
    for wcol in range(weeks):
        mon = w0 + datetime.timedelta(weeks=wcol)
        if it[5] <= mon + datetime.timedelta(days=6) and it[6] >= mon:
            ws.cell(r, 7 + wcol).fill = bar_fill

for mname, mdate in milestones.items():                   # вертикальные маркеры вех
    if mdate is None:
        continue
    col = 7 + max(0, (mdate - w0).days // 7)
    ws.cell(1, col, mname).font = Font(bold=True, size=8)
    for r in range(2, 4 + len(items)):
        ws.cell(r, col).fill = mile_fill

ws.column_dimensions['A'].width = 6
ws.column_dimensions['B'].width = 60
for cl in 'CDEF':
    ws.column_dimensions[cl].width = 13
for wcol in range(weeks):
    ws.column_dimensions[get_column_letter(7 + wcol)].width = 2.2
ws.freeze_panes = 'G4'
ws.page_setup.orientation = 'landscape'
ws.page_setup.paperSize = 8                               # A3
ws.sheet_properties.pageSetUpPr.fitToPage = True
ws.page_setup.fitToWidth = 1
ws.page_setup.fitToHeight = 1

# ---------------- 7. Выдача ----------------
buf = io.BytesIO()
out.save(buf)
buf.seek(0)

st.success(f'✅ Готово: {len(items)} строк, финиш графика: {items[-1][6]:%d.%m.%Y}')
st.download_button('⬇️ Скачать SMS_P25077.xlsx (таблица + диаграмма Ганта, 1 лист A3)',
                   buf, file_name='SMS_P25077.xlsx',
                   mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
