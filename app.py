# pip install openpyxl
import openpyxl, datetime
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from openpyxl.utils import get_column_letter

SRC = 'PLANT_MASTER_SCHEDULE P25077.xlsx'
OUT = 'SMS_P25077.xlsx'
PHASE2_START = datetime.date(2025, 6, 11)          # старт Фазы 2 (A13)
MILESTONES = {'VC': datetime.date(2028,9,24), 'PT1': datetime.date(2029,1,14),
              'PT2': datetime.date(2030,4,22), 'PPC': datetime.date(2030,8,19),
              'PSW': datetime.date(2031,10,19), 'SOP': datetime.date(2031,11,30)}

# --- Праздники РФ (производственный календарь, сверить переносы на год выпуска) ---
HOL = set()
for y in range(2025, 2033):
    HOL |= {datetime.date(y,1,d) for d in range(1,9)}
    HOL |= {datetime.date(y,2,23), datetime.date(y,3,8), datetime.date(y,5,1),
            datetime.date(y,5,2), datetime.date(y,5,9), datetime.date(y,6,12),
            datetime.date(y,11,4), datetime.date(y,12,31)}
HOL |= {datetime.date(2025,3,9), datetime.date(2026,3,9)}  # 8 марта в сб/вс
def is_work(d): return d.weekday() < 5 and d not in HOL
def next_work(d):
    while not is_work(d): d += datetime.timedelta(days=1)
    return d
def add_workdays(start, n):          # n рабочих дней ВКЛЮЧИТЕЛЬНО
    d, c = start, 1
    while c < n:
        d += datetime.timedelta(days=1)
        if is_work(d): c += 1
    return d

# --- 1. Чтение данных ---
wbv = openpyxl.load_workbook(SRC, data_only=True)   # кэш значений формул
CFG = [('PMSPR TOGF-ENG-008-06 Phase2', 36, 'Фаза 2 — Разработка продукта и процесса'),
       ('PMSPR TOGF-ENG-008-06 Phase 3', 39, 'Фаза 3 — Индустриализация'),
       ('PMSPR TOGF-ENG-008-06 Phase4;5', 39, 'Фаза 4/5 — Наращивание / Серийная жизнь')]
items = []
for name, resp_col, phase in CFG:
    ws = wbv[name]
    for r in range(13, ws.max_row + 1):
        desc = ws.cell(r, 25).value
        if not (isinstance(desc, str) and desc.strip()): continue
        C, L = ws.cell(r, 3).value, ws.cell(r, 12).value   # item start / item end
        dur = int(round(L - C + 1)) if isinstance(C,(int,float)) and isinstance(L,(int,float)) and L >= C else None
        items.append([str(ws.cell(r,24).value or '').strip(), desc.strip(),
                      ws.cell(r, resp_col).value or '', phase, dur])
missing = sum(1 for i in items if i[4] is None)
assert missing == 0, f'{missing} строк без кэша формул! Откройте файл в Excel, сохраните, повторите.'
# ИГНОРИРУЮТСЯ: PLANNED START/END DATE, START/END WEEK, ACTUAL END DATE STATUS

# --- 2. Расчёт календаря (последовательная цепочка, только рабочие дни РФ) ---
cur = next_work(PHASE2_START)
for it in items:
    wd = max(1, int(round((it[4] or 0.3) * 5)))
    it += [cur, add_workdays(cur, wd)]                 # Start, Finish
    cur = next_work(it[-1] + datetime.timedelta(days=1))

# --- 3. Сборка SMS: таблица + диаграмма Ганта на ОДНОМ листе ---
out = openpyxl.Workbook(); ws = out.active; ws.title = 'SMS P25077'
hdr = ['№', 'Description', 'Phase', 'Duration, weeks', 'Start', 'Finish']
for c, h in enumerate(hdr, 1):
    cell = ws.cell(3, c, h); cell.font = Font(bold=True, color='FFFFFF')
    cell.fill = PatternFill('solid', fgColor='4472C4')
    cell.alignment = Alignment(horizontal='center', wrap_text=True)
bar = PatternFill('solid', fgColor='2E75B6'); mile_f = PatternFill('solid', fgColor='FFC000')
w0 = items[0][5] - datetime.timedelta(days=items[0][5].weekday())   # понедельник стартовой недели
wk_end = max(i[6] for i in items)
weeks = ((wk_end - w0).days // 7) + 2
for wcol in range(weeks):                            # шапка недель
    mon = w0 + datetime.timedelta(weeks=wcol)
    c = ws.cell(2, 7 + wcol, f'CW{mon.isocalendar().week}/{mon.year}')
    c.font = Font(size=7); c.alignment = Alignment(horizontal='center', textRotation=90)
thin = Border(*[Side(style='thin', color='D9D9D9')]*4)
for i, it in enumerate(items):
    r = 4 + i
    for c, v in enumerate(it[:6], 1):
        cell = ws.cell(r, c, v); cell.border = thin
        if c in (5, 6): cell.number_format = 'DD.MM.YYYY'
    ws.cell(r, 2).alignment = Alignment(wrap_text=True, vertical='center')
    for wcol in range(weeks):                        # полоса Ганта
        mon = w0 + datetime.timedelta(weeks=wcol)
        sun = mon + datetime.timedelta(days=6)
        if it[5] <= sun and it[6] >= mon:
            ws.cell(r, 7 + wcol).fill = bar
for mname, mdate in MILESTONES.items():              # вертикальные маркеры вех
    wcol = max(0, ((mdate - w0).days // 7))
    col = 7 + wcol
    for r in range(2, 4 + len(items)):
        ws.cell(r, col).fill = mile_f
    ws.cell(1, col, mname).font = Font(bold=True, size=8)
ws.column_dimensions['A'].width = 6; ws.column_dimensions['B'].width = 60
for cl in 'CDE': ws.column_dimensions[cl].width = 13
for wcol in range(weeks): ws.column_dimensions[get_column_letter(7+wcol)].width = 2.2
ws.freeze_panes = 'G4'
ws.page_setup.orientation = 'landscape'; ws.page_setup.paperSize = 8  # A3
ws.page_setup.fitToWidth = 1; ws.page_setup.fitToHeight = 1
ws.sheet_properties.pageSetUpPr.fitToPage = True
out.save(OUT)
print(f'OK: {OUT}, строк: {len(items)}, финиш графика: {items[-1][6]}')
