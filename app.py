# -*- coding: utf-8 -*-
import streamlit as st
import openpyxl
import datetime
import io
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

# -------- Настройки --------
PHASE2_START = datetime.date(2025, 6, 11)

MILESTONES = {
    'VC':  datetime.date(2028, 9, 24),
    'PT1': datetime.date(2029, 1, 14),
    'PT2': datetime.date(2030, 4, 22),
    'PPC': datetime.date(2030, 8, 19),
    'PSW': datetime.date(2031, 10, 19),
    'SOP': datetime.date(2031, 11, 30),
}

CFG = [
    ('PMSPR TOGF-ENG-008-06 Phase2',   36, 'Фаза 2 — Разработка продукта и процесса'),
    ('PMSPR TOGF-ENG-008-06 Phase 3',  39, 'Фаза 3 — Индустриализация'),
    ('PMSPR TOGF-ENG-008-06 Phase4;5', 39, 'Фаза 4/5 — Наращивание / Серийная жизнь'),
]

# -------- Праздники РФ --------
HOL = set()
for y in range(2025, 2033):
    HOL |= {datetime.date(y, 1, d) for d in range(1, 9)}
    HOL |= {
        datetime.date(y, 2, 23), datetime.date(y, 3, 8),
        datetime.date(y, 5, 1),  datetime.date(y, 5, 2),
        datetime.date(y, 5, 9),  datetime.date(y, 6, 12),
        datetime.date(y, 11, 4), datetime.date(y, 12, 31),
    }
HOL |= {datetime.date(2025, 3, 9), datetime.date(2026, 3, 9)}


def is_work(d):
    return d.weekday() < 5 and d not in HOL


def next_work(d):
    while not is_work(d):
        d += datetime.timedelta(days=1)
    return d


def add_workdays(start, n):
    d, c = start, 1
    while c < n:
        d += datetime.timedelta(days=1)
        if is_work(d):
            c += 1
    return d


# -------- Логика обработки --------
def parse_source(wbv):
    items = []
    for name, resp_col, phase in CFG:
        if name not in wbv.sheetnames:
            raise ValueError("Вкладка не найдена: " + name)
        ws = wbv[name]
        for r in range(13, ws.max_row + 1):
            desc = ws.cell(r, 25).value
            if not (isinstance(desc, str) and desc.strip()):
                continue
            C = ws.cell(r, 3).value
            L = ws.cell(r, 12).value
            if isinstance(C, (int, float)) and isinstance(L, (int, float)) and L >= C:
                dur = int(round(L - C + 1))
            else:
                dur = None
            items.append([
                str(ws.cell(r, 24).value or '').strip(),
                desc.strip(),
                ws.cell(r, resp_col).value or '',
                phase,
                dur,
            ])
    missing = sum(1 for i in items if i[4] is None)
    if missing:
        raise ValueError(
            str(missing) + " строк без кэша формул. Откройте файл в Excel, "
            "сохраните и повторите загрузку."
        )
    return items


def build_schedule(items):
    cur = next_work(PHASE2_START)
    for it in items:
        wd = max(1, int(round((it[4] or 0.3) * 5)))
        it.append(cur)
        it.append(add_workdays(cur, wd))
        cur = next_work(it[-1] + datetime.timedelta(days=1))
    return items


def build_workbook(items):
    out = openpyxl.Workbook()
    ws = out.active
    ws.title = 'SMS P25077'

    hdr = ['№', 'Description', 'Phase', 'Duration, weeks', 'Start', 'Finish']
    for c, h in enumerate(hdr, 1):
        cell = ws.cell(3, c, h)
        cell.font = Font(bold=True, color='FFFFFF')
        cell.fill = PatternFill('solid', fgColor='4472C4')
        cell.alignment = Alignment(horizontal='center', wrap_text=True)

    bar = PatternFill('solid', fgColor='2E75B6')
    mile_f = PatternFill('solid', fgColor='FFC000')

    w0 = items[0][5] - datetime.timedelta(days=items[0][5].weekday())
    wk_end = max(i[6] for i in items)
    weeks = ((wk_end - w0).days // 7) + 2

    for wcol in range(weeks):
        mon = w0 + datetime.timedelta(weeks=wcol)
        c = ws.cell(2, 7 + wcol, 'CW' + str(mon.isocalendar().week) + '/' + str(mon.year))
        c.font = Font(size=7)
        c.alignment = Alignment(horizontal='center', textRotation=90)

    thin = Border(*[Side(style='thin', color='D9D9D9')] * 4)

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
            sun = mon + datetime.timedelta(days=6)
            if it[5] <= sun and it[6] >= mon:
                ws.cell(r, 7 + wcol).fill = bar

    for mname, mdate in MILESTONES.items():
        wcol = max(0, (mdate - w0).days // 7)
        col = 7 + wcol
        for r in range(2, 4 + len(items)):
            ws.cell(r, col).fill = mile_f
        ws.cell(1, col, mname).font = Font(bold=True, size=8)

    ws.column_dimensions['A'].width = 6
    ws.column_dimensions['B'].width = 60
    for cl in 'CDE':
        ws.column_dimensions[cl].width = 13
    for wcol in range(weeks):
        ws.column_dimensions[get_column_letter(7 + wcol)].width = 2.2

    ws.freeze_panes = 'G4'
    ws.page_setup.orientation = 'landscape'
    ws.page_setup.paperSize = 8
    ws.page_setup.fitToWidth = 1
    ws.page_setup.fitToHeight = 1
    ws.sheet_properties.pageSetUpPr.fitToPage = True

    stream = io.BytesIO()
    out.save(stream)
    stream.seek(0)
    return stream


# -------- Streamlit интерфейс --------
st.set_page_config(page_title='SMS P25077', layout='wide')
st.title('SMS график проекта P25077')
st.write('Загрузите PLANT_MASTER_SCHEDULE P25077.xlsx для формирования графика.')

uploaded = st.file_uploader('Excel-файл', type=['xlsx', 'xlsm'])

if uploaded is not None:
    if st.button('Сформировать график'):
        try:
            with st.spinner('Чтение файла...'):
                wbv = openpyxl.load_workbook(uploaded, data_only=True)
                items = parse_source(wbv)
                items = build_schedule(items)

            st.success('Готово. Строк: ' + str(len(items)) +
                       '. Финиш графика: ' + str(items[-1][6]))

            with st.spinner('Формирование Excel...'):
                stream = build_workbook(items)

            st.download_button(
                label='Скачать SMS_P25077.xlsx',
                data=stream,
                file_name='SMS_P25077.xlsx',
                mime='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet',
            )

            st.subheader('Предпросмотр')
            preview = [
                {
                    '№': i + 1,
                    'Description': it[1],
                    'Phase': it[3],
                    'Duration, weeks': it[4],
                    'Start': it[5].strftime('%d.%m.%Y'),
                    'Finish': it[6].strftime('%d.%m.%Y'),
                }
                for i, it in enumerate(items)
            ]
            st.dataframe(preview, use_container_width=True)

        except Exception as e:
            st.error('Ошибка: ' + str(e))
