"""
DDS Allada  —  Сравнение расходов по периодам.

Flask-приложение: выбор периода (месяц / квартал / год),
таблица сравнения год-к-году в браузере + скачивание XLSX.

Выручка берётся из строки 18 листа ДДС
  «Поступления от клиентов (продажа товаров/услуг)»
Кредитные операции выносятся в отдельный блок внизу.

python app.py  →  http://localhost:5000
"""
from flask import Flask, render_template, request, send_file, jsonify
import pandas as pd
import openpyxl
from openpyxl.styles import PatternFill, Font, Alignment, Border, Side
from collections import defaultdict
from pathlib import Path
from werkzeug.utils import secure_filename
from datetime import datetime
import calendar, io, warnings, re
warnings.filterwarnings('ignore')

app = Flask(__name__)


@app.after_request
def add_cors_headers(resp):
    """
    Нужен для сценария, когда UI открыт как file://.../index.html
    и обращается к API на http://127.0.0.1:5000.
    """
    if request.path.startswith('/api/'):
        resp.headers['Access-Control-Allow-Origin'] = '*'
        resp.headers['Access-Control-Allow-Methods'] = 'GET, POST, OPTIONS'
        resp.headers['Access-Control-Allow-Headers'] = 'Content-Type'
    return resp

# ── Config ─────────────────────────────────────────────────────────────────────
XLSX_IN = Path(__file__).resolve().parent / "system_dds_allada.xlsx"
TARGET_FILE_NAME = "system_dds_allada.xlsx"
ALLOWED_EXTENSIONS = {".xlsx"}

BANK_WALLETS = {'Альфа 796', 'Сбер 595', 'Счет 606', 'Счет 11', 'Счет 12'}
EXCLUDE_ARTICLES = {
    'Доход — Перевод между счетами',  'Расход — Перевод между счетами',
    'Доход — перевод между счетами',  'Расход — перевод между счетами',
}
RENAME_ARTICLES = {'Реклама/Маркетинг Звягин': 'Привлечение по ДМС Звягин'}

# Статьи, относящиеся к кредитам/займам (не расходы и не доходы)
CREDIT_ARTICLES_SET = {
    'Возврат кредитов и займов',
    'Получение кредитов и займов',
    'Полученные проценты за выданные кредиты и займы',
    'Проценты по кредиту',
}

JOURNAL_SHEETS = [
    ('Журнал Снабжение',   0, 2, 6, 1),
    ('Журнал МобиДик',     0, 2, 6, 1),
    ('Журнал Леонова',     0, 2, 6, 1),
    ('Журнал Мира 100',    0, 2, 6, 1),
    ('Журнал Таштемирова', 0, 2, 6, 1),
    ('Журнал Ирина Влад.', 0, 2, 6, 1),
]
BANK_SHEETS = [('1С Альфа', 0, 12, 13, 0), ('1С Сбер', 0, 12, 13, 0)]

MONTH_RU = {1:'Январь',2:'Февраль',3:'Март',4:'Апрель',5:'Май',6:'Июнь',
            7:'Июль',8:'Август',9:'Сентябрь',10:'Октябрь',11:'Ноябрь',12:'Декабрь'}

# ── Маппинг колонок листа ДДС ─────────────────────────────────────────────────
# ДДС row 18 (0-indexed 17) = "Поступления от клиентов (продажа товаров/услуг):"
# Pandas col offsets (0-indexed) per year:
#   2024: month M → col M        (Jan=1, Dec=12)
#   2025: month M → col 14 + M   (Jan=15, Dec=26)
#   2026: month M → col 31 + M   (Jan=32, Dec=43)
DDS_REVENUE_ROW = 17    # pandas 0-indexed (Excel row 18)
DDS_YEAR_OFFSETS = {2024: 0, 2025: 14, 2026: 31}

# ── Data loading (cached) ─────────────────────────────────────────────────────
_TX = None
_DDS = None


def reset_cache():
    global _TX, _DDS
    _TX = None
    _DDS = None


def get_current_file_info():
    return {
        'path': str(XLSX_IN),
        'name': XLSX_IN.name,
        'exists': XLSX_IN.exists(),
    }


def _load_sheet(name, dc, ac, rc, hr, wc=None):
    try:
        df = pd.read_excel(XLSX_IN, sheet_name=name, header=hr, dtype=str)
    except Exception:
        return pd.DataFrame()
    cols = df.columns.tolist()
    if max(dc, ac, rc) >= len(cols):
        return pd.DataFrame()
    df = df.rename(columns={cols[dc]: 'date', cols[ac]: 'amount', cols[rc]: 'article'})
    keep = ['date', 'amount', 'article']
    if wc is not None and wc < len(cols):
        df = df.rename(columns={cols[wc]: 'wallet'}); keep.append('wallet')
    df = df[keep].copy()
    df['date'] = pd.to_datetime(df['date'], errors='coerce')
    df['amount'] = pd.to_numeric(df['amount'], errors='coerce')
    df = df.dropna(subset=['date', 'amount', 'article'])
    df = df[df['article'].str.strip().ne('')]
    if 'wallet' in df.columns:
        df = df[~df['wallet'].isin(BANK_WALLETS)].drop(columns=['wallet'])
    return df


def get_tx():
    global _TX
    if _TX is None:
        if not XLSX_IN.exists():
            raise FileNotFoundError(f'Файл не найден: {XLSX_IN}')
        print("Загрузка транзакций из Excel...")
        frames = [_load_sheet(*c, wc=1) for c in JOURNAL_SHEETS] + \
                 [_load_sheet(*c) for c in BANK_SHEETS]
        _TX = pd.concat([f for f in frames if not f.empty], ignore_index=True)
        _TX['article'] = _TX['article'].str.strip().replace(RENAME_ARTICLES)
        print(f"  Транзакций: {len(_TX):,}")
    return _TX


def get_dds():
    """Кэшированное чтение листа ДДС (числа, без заголовков)."""
    global _DDS
    if _DDS is None:
        if not XLSX_IN.exists():
            raise FileNotFoundError(f'Файл не найден: {XLSX_IN}')
        print("Загрузка листа ДДС...")
        _DDS = pd.read_excel(XLSX_IN, sheet_name='ДДС', header=None)
    return _DDS


# ── Revenue from ДДС sheet ─────────────────────────────────────────────────────
def get_revenue(year, months):
    """
    Выручка из строки 18 листа ДДС
    «Поступления от клиентов (продажа товаров/услуг)»
    Суммируем значения по указанным месяцам для данного года.
    """
    dds = get_dds()
    offset = DDS_YEAR_OFFSETS.get(year)
    if offset is None:
        return 0.0
    total = 0.0
    for m in months:
        col_idx = offset + m
        if col_idx < dds.shape[1]:
            v = dds.iloc[DDS_REVENUE_ROW, col_idx]
            try:
                total += float(v)
            except (ValueError, TypeError):
                pass
    return total


# ── Core analytics ─────────────────────────────────────────────────────────────
def is_credit_article(name):
    return name in CREDIT_ARTICLES_SET


def build_pivot(tx, date_from, date_to, months, credit=False):
    """
    credit=False → обычные расходы (без кредитов, amount < 0)
    credit=True  → все кредитные операции (amount < 0 И amount > 0)
    """
    base = tx[(tx['date'] >= date_from) & (tx['date'] <= date_to) &
              (~tx['article'].isin(EXCLUDE_ARTICLES))].copy()

    if credit:
        base = base[base['article'].apply(is_credit_article)]
        # Для кредитов: сохраняем знак (выдача = расход < 0, возврат = доход > 0)
        base['val'] = base['amount']
    else:
        base = base[(base['amount'] < 0) & (~base['article'].apply(is_credit_article))]
        base['val'] = base['amount'].abs()

    base['mon'] = base['date'].dt.month
    pv = base.groupby(['article', 'mon'])['val'].sum().unstack(fill_value=0)
    for m in months:
        if m not in pv.columns:
            pv[m] = 0
    return pv[months]


def detect_groups(articles):
    pfx = defaultdict(list)
    for a in articles:
        pfx[a.split()[0].rstrip('/')].append(a)
    out = {}
    for k, arts in pfx.items():
        lbl = k.upper() if len(arts) >= 2 else arts[0]
        for a in arts:
            out[a] = lbl
    return out


def get_date_range(period_type, period_value, year):
    y = int(year)
    if period_type == 'month':
        m = int(period_value)
        last = calendar.monthrange(y, m)[1]
        return f'{y}-{m:02d}-01', f'{y}-{m:02d}-{last:02d}', [m]
    elif period_type == 'quarter':
        q = int(period_value)
        ms = {1: [1, 2, 3], 2: [4, 5, 6], 3: [7, 8, 9], 4: [10, 11, 12]}[q]
        last = calendar.monthrange(y, ms[-1])[1]
        return f'{y}-{ms[0]:02d}-01', f'{y}-{ms[-1]:02d}-{last:02d}', ms
    else:
        return f'{y}-01-01', f'{y}-12-31', list(range(1, 13))


def _build_rows(pv_new, pv_old, months, rev_old, rev_new):
    """Построить список строк из двух пивотов. Возвращает (rows, total_old, total_new)."""
    all_arts = list(dict.fromkeys(
        list(pv_new.index) + [a for a in pv_old.index if a not in pv_new.index]
    ))
    if not all_arts:
        return [], 0, 0

    art_to_grp = detect_groups(all_arts)
    gmap = defaultdict(list)
    for a, g in art_to_grp.items():
        gmap[g].append(a)
    order = sorted(gmap, key=lambda g:
        abs(pv_new.loc[pv_new.index.isin(gmap[g])].sum().sum())
        if not pv_new.loc[pv_new.index.isin(gmap[g])].empty else 0,
        reverse=True)

    def safe_pct(v, base):
        return round(v / base * 100, 1) if base else None

    rows = []
    g_old_total = g_new_total = 0

    for grp in order:
        arts = gmap[grp]
        g_new_vals = [float(pv_new.loc[pv_new.index.isin(arts), m].sum()) for m in months]
        g_old_vals = [float(pv_old.loc[pv_old.index.isin(arts), m].sum()) for m in months]
        g_new_sum = sum(g_new_vals)
        g_old_sum = sum(g_old_vals)
        g_old_total += g_old_sum
        g_new_total += g_new_sum

        delta = g_new_sum - g_old_sum
        pct = round(delta / g_old_sum * 100, 1) if g_old_sum else (100 if g_new_sum else 0)
        pct_ro = safe_pct(g_old_sum, rev_old)
        pct_rn = safe_pct(g_new_sum, rev_new)
        dpp = round(pct_rn - pct_ro, 1) if pct_ro is not None and pct_rn is not None else None

        if g_old_sum == 0 and g_new_sum == 0:    assess = '—'
        elif g_old_sum == 0:                      assess = 'новая статья'
        elif g_new_sum == 0:                      assess = 'статья закрыта'
        elif pct > 20:                            assess = f'рост {abs(pct):.0f}%'
        elif pct < -20:                           assess = f'снижение {abs(pct):.0f}%'
        else:                                     assess = 'стабильно'

        rows.append({
            'type': 'group', 'label': grp, 'single': len(arts) == 1,
            'months_new': [round(v) for v in g_new_vals],
            'months_old': [round(v) for v in g_old_vals],
            'avg_new': round(g_new_sum / len(months)),
            'total_new': round(g_new_sum), 'total_old': round(g_old_sum),
            'delta': round(delta), 'pct': pct, 'assess': assess,
            'pct_rev_old': pct_ro, 'pct_rev_new': pct_rn, 'delta_pp': dpp,
        })

        if len(arts) > 1:
            for art in sorted(arts,
                    key=lambda a: abs(pv_new.loc[a, months].sum()) if a in pv_new.index else 0,
                    reverse=True):
                a_new = [float(pv_new.loc[art, m]) if art in pv_new.index else 0 for m in months]
                a_old = [float(pv_old.loc[art, m]) if art in pv_old.index else 0 for m in months]
                an = sum(a_new); ao = sum(a_old); d = an - ao
                p = round(d / ao * 100, 1) if ao else (100 if an else 0)
                rows.append({
                    'type': 'detail', 'label': art,
                    'months_new': [round(v) for v in a_new],
                    'months_old': [round(v) for v in a_old],
                    'avg_new': round(an / len(months)),
                    'total_new': round(an), 'total_old': round(ao),
                    'delta': round(d), 'pct': p, 'assess': '',
                    'pct_rev_old': safe_pct(ao, rev_old),
                    'pct_rev_new': safe_pct(an, rev_new),
                    'delta_pp': None,
                })

    return rows, g_old_total, g_new_total


def build_report(period_type, period_value, year_new, year_old):
    tx = get_tx()
    yn, yo = int(year_new), int(year_old)
    new_from, new_to, months = get_date_range(period_type, period_value, yn)
    old_from, old_to, _      = get_date_range(period_type, period_value, yo)

    # Выручка из листа ДДС строка 18
    rev_new = get_revenue(yn, months)
    rev_old = get_revenue(yo, months)

    # ── Расходы (без кредитов) ────────────────────────────────────────────────
    pv_new = build_pivot(tx, new_from, new_to, months, credit=False)
    pv_old = build_pivot(tx, old_from, old_to, months, credit=False)
    pv_new = pv_new[(pv_new > 0).any(axis=1)]

    expense_rows, exp_old, exp_new = _build_rows(pv_new, pv_old, months, rev_old, rev_new)

    # Grand total расходы
    gd = exp_new - exp_old
    gp = round(gd / exp_old * 100, 1) if exp_old else 0
    def sp(v, b): return round(v / b * 100, 1) if b else None
    expense_rows.append({
        'type': 'grand', 'label': 'ИТОГО РАСХОДЫ',
        'months_new': [], 'months_old': [],
        'avg_new': round(exp_new / len(months)),
        'total_new': round(exp_new), 'total_old': round(exp_old),
        'delta': round(gd), 'pct': gp, 'assess': '',
        'pct_rev_old': sp(exp_old, rev_old), 'pct_rev_new': sp(exp_new, rev_new),
        'delta_pp': round((sp(exp_new, rev_new) or 0) - (sp(exp_old, rev_old) or 0), 1)
                   if rev_old and rev_new else None,
    })

    # ── Кредиты и займы (отдельный блок) ──────────────────────────────────────
    cr_new = build_pivot(tx, new_from, new_to, months, credit=True)
    cr_old = build_pivot(tx, old_from, old_to, months, credit=True)
    credit_rows, cr_old_t, cr_new_t = _build_rows(cr_new, cr_old, months, rev_old, rev_new)

    if credit_rows:
        crd = cr_new_t - cr_old_t
        credit_rows.append({
            'type': 'grand', 'label': 'ИТОГО КРЕДИТЫ И ЗАЙМЫ',
            'months_new': [], 'months_old': [],
            'avg_new': round(cr_new_t / len(months)),
            'total_new': round(cr_new_t), 'total_old': round(cr_old_t),
            'delta': round(crd),
            'pct': round(crd / cr_old_t * 100, 1) if cr_old_t else 0,
            'assess': '', 'pct_rev_old': None, 'pct_rev_new': None, 'delta_pp': None,
        })

    # Period label
    if period_type == 'month':    plabel = MONTH_RU[int(period_value)]
    elif period_type == 'quarter': plabel = f'Q{period_value}'
    else:                          plabel = 'Год'

    # Помесячная выручка
    rev_new_months = [round(get_revenue(yn, [m])) for m in months]
    rev_old_months = [round(get_revenue(yo, [m])) for m in months]

    rev_delta = round(rev_new) - round(rev_old)
    rev_pct = round(rev_delta / rev_old * 100, 1) if rev_old else 0

    return {
        'period_label': plabel,
        'period_type': period_type,
        'period_value': period_value,
        'year_new': yn, 'year_old': yo,
        'months': months,
        'month_names': [MONTH_RU[m] for m in months],
        'rev_new': round(rev_new),
        'rev_old': round(rev_old),
        'rev_new_months': rev_new_months,
        'rev_old_months': rev_old_months,
        'rev_delta': rev_delta,
        'rev_pct': rev_pct,
        'rows': expense_rows,
        'credit_rows': credit_rows,
    }


# ── XLSX generation ────────────────────────────────────────────────────────────
def _fill(h): return PatternFill('solid', fgColor=h)
def _fnt(bold=False, color='000000', size=10, italic=False):
    return Font(bold=bold, color=color, size=size, italic=italic)
def _aln(h='right', v='center'): return Alignment(horizontal=h, vertical=v)
_s = Side(style='thin', color='B0BEC5')
_BRD = Border(left=_s, right=_s, top=_s, bottom=_s)
_NUM = '#,##0'


def _wcell(ws, r, c, val, fl, fn, fmt=None, ha='right'):
    cell = ws.cell(r, c, val)
    cell.fill = fl; cell.font = fn; cell.border = _BRD; cell.alignment = _aln(ha)
    if fmt: cell.number_format = fmt
    return cell


def _row_style(row):
    t = row['type']
    if t == 'grand':
        return _fill('0D1E35'), _fnt(bold=True, color='FFFFFF', size=11)
    if t == 'detail':
        return _fill('FAFBFD'), _fnt(italic=True, color='444466', size=9)
    if row.get('single'):
        return _fill('CFD8E8'), _fnt(bold=True, color='1A1A2E', size=10)
    return _fill('1A3A5C'), _fnt(bold=True, color='FFFFFF', size=10)


def _credit_style(row):
    t = row['type']
    if t == 'grand':
        return _fill('3d2c1e'), _fnt(bold=True, color='FFFFFF', size=11)
    if t == 'detail':
        return _fill('FFF8F0'), _fnt(italic=True, color='7d5a3a', size=9)
    if row.get('single'):
        return _fill('F5E6D0'), _fnt(bold=True, color='5a3a1a', size=10)
    return _fill('8B5E3C'), _fnt(bold=True, color='FFFFFF', size=10)


def _write_block(ws, start_row, all_rows, months, mnames, yn, style_fn):
    """Write a block of rows to a data sheet. Returns next row."""
    n = len(months); r = start_row
    for row in all_rows:
        fl, fn = style_fn(row)
        label = ('\u00a0' * 4 + 'в т.ч. ' + row['label']) if row['type'] == 'detail' else row['label']
        _wcell(ws, r, 1, label, fl, fn, ha='left')
        for i, v in enumerate(row.get('months_new', [])):
            _wcell(ws, r, i + 2, v, fl, fn, _NUM)
        if not row.get('months_new'):
            for i in range(n):
                _wcell(ws, r, i + 2, '', fl, fn)
        _wcell(ws, r, n + 2, row['avg_new'], fl, fn, _NUM)
        _wcell(ws, r, n + 3, row['total_new'], fl, fn, _NUM)
        r += 1
    return r


def _write_cmp_block(ws, start_row, all_rows, style_fn):
    """Write a block of rows to a comparison sheet. Returns next row."""
    r = start_row
    for row in all_rows:
        fl, fn = style_fn(row)
        label = ('\u00a0' * 4 + 'в т.ч. ' + row['label']) if row['type'] == 'detail' else row['label']
        _wcell(ws, r, 1, label, fl, fn, ha='left')
        _wcell(ws, r, 2, row['total_old'], fl, fn, _NUM)
        _wcell(ws, r, 3, row['total_new'], fl, fn, _NUM)
        _wcell(ws, r, 4, row['delta'], fl, fn, '+#,##0;-#,##0;—')
        _wcell(ws, r, 5, f"{row['pct']:+.1f}%", fl, fn)
        _wcell(ws, r, 6, row.get('assess', ''), fl, fn, ha='center')
        pro = row.get('pct_rev_old')
        prn = row.get('pct_rev_new')
        _wcell(ws, r, 7, f"{pro:.1f}%" if pro is not None else '—', fl, fn)
        _wcell(ws, r, 8, f"{prn:.1f}%" if prn is not None else '—', fl, fn)
        dpp = row.get('delta_pp')
        _wcell(ws, r, 9, f"{dpp:+.1f} пп" if dpp is not None else '—', fl, fn, ha='center')
        r += 1
    return r


def generate_xlsx(report):
    wb = openpyxl.Workbook()
    rows = report['rows']
    credit = report.get('credit_rows', [])
    months = report['months']; mnames = report['month_names']
    yn = report['year_new']; yo = report['year_old']
    n = len(months)

    # ── Sheet 1: Data ───────────────────────────────────────────────────────
    ws1 = wb.active
    ws1.title = f"{report['period_label']} {yn}"
    ws1.column_dimensions['A'].width = 50
    for i in range(1, n + 3):
        ws1.column_dimensions[chr(ord('A') + i)].width = 15

    last_col = chr(ord('A') + n + 2)
    ws1.merge_cells(f'A1:{last_col}1')
    ws1['A1'].value = f"Расходы — {report['period_label']} {yn}"
    ws1['A1'].font = _fnt(bold=True, color='FFFFFF', size=12)
    ws1['A1'].fill = _fill('1A3A5C'); ws1['A1'].alignment = _aln('center')

    hdrs = ['Статья / Группа'] + [f'{m} {yn}' for m in mnames] + ['Среднее/мес', 'Итого']
    for ci, h in enumerate(hdrs, 1):
        c = ws1.cell(2, ci, h)
        c.font = _fnt(bold=True, color='FFFFFF', size=9)
        c.fill = _fill('1A3A5C'); c.border = _BRD
        c.alignment = _aln('center' if ci == 1 else 'right')

    r = _write_block(ws1, 3, rows, months, mnames, yn, _row_style)

    if credit:
        r += 1  # blank row
        sec = ws1.cell(r, 1, 'КРЕДИТЫ И ЗАЙМЫ')
        sec.font = _fnt(bold=True, color='8B5E3C', size=11)
        sec.fill = _fill('F5E6D0'); sec.border = _BRD
        ws1.merge_cells(f'A{r}:{last_col}{r}')
        sec.alignment = _aln('center')
        r += 1
        _write_block(ws1, r, credit, months, mnames, yn, _credit_style)

    ws1.freeze_panes = 'B3'

    # ── Sheet 2: Comparison ─────────────────────────────────────────────────
    ws2 = wb.create_sheet(f"vs {yo}")
    ws2.column_dimensions['A'].width = 50
    for col in list('BCDEFGHI'):
        ws2.column_dimensions[col].width = 15

    ws2.merge_cells('A1:I1')
    ws2['A1'].value = (f"Сравнение {report['period_label']}: {yn} vs {yo}   "
                       f"Выручка {yo}: {report['rev_old']:,} | "
                       f"Выручка {yn}: {report['rev_new']:,}")
    ws2['A1'].font = _fnt(bold=True, color='FFFFFF', size=11)
    ws2['A1'].fill = _fill('1A3A5C'); ws2['A1'].alignment = _aln('center')

    hdrs2 = ['Статья / Группа', f'{yo}', f'{yn}',
             'Изменение', 'Тренд %', 'Оценка',
             f'% выр. {yo}', f'% выр. {yn}', 'Д пп']
    for ci, h in enumerate(hdrs2, 1):
        c = ws2.cell(2, ci, h)
        c.font = _fnt(bold=True, color='FFFFFF', size=9)
        c.fill = _fill('1A3A5C'); c.border = _BRD
        c.alignment = _aln('center' if ci == 1 else 'right')

    r = _write_cmp_block(ws2, 3, rows, _row_style)

    if credit:
        r += 1
        sec = ws2.cell(r, 1, 'КРЕДИТЫ И ЗАЙМЫ')
        sec.font = _fnt(bold=True, color='8B5E3C', size=11)
        sec.fill = _fill('F5E6D0'); sec.border = _BRD
        ws2.merge_cells(f'A{r}:I{r}')
        sec.alignment = _aln('center')
        r += 1
        _write_cmp_block(ws2, r, credit, _credit_style)

    ws2.freeze_panes = 'B3'

    buf = io.BytesIO()
    wb.save(buf); buf.seek(0)
    return buf


# ── Routes ─────────────────────────────────────────────────────────────────────
@app.route('/')
def index():
    return render_template('index.html', file_info=get_current_file_info())


@app.route('/health', methods=['GET'])
def health():
    return jsonify({
        'ok': True,
        'service': 'dds_allada',
        'file': get_current_file_info(),
    })


@app.route('/api/file-info', methods=['GET'])
def api_file_info():
    return jsonify(get_current_file_info())


@app.route('/api/upload', methods=['POST'])
def api_upload():
    global XLSX_IN
    try:
        if 'file' not in request.files:
            return jsonify({'ok': False, 'error': 'Файл не передан.'}), 400

        uploaded = request.files['file']
        if not uploaded or not uploaded.filename:
            return jsonify({'ok': False, 'error': 'Выберите файл для загрузки.'}), 400

        safe_name = secure_filename(uploaded.filename)
        ext = Path(safe_name).suffix.lower()
        if ext not in ALLOWED_EXTENSIONS:
            return jsonify({'ok': False, 'error': 'Разрешен только формат .xlsx.'}), 400

        upload_dir = Path(__file__).resolve().parent / 'uploads'
        upload_dir.mkdir(parents=True, exist_ok=True)
        target = upload_dir / TARGET_FILE_NAME
        uploaded.save(target)

        # Проверяем, что файл действительно читается как Excel workbook.
        openpyxl.load_workbook(target, read_only=True).close()

        XLSX_IN = target
        reset_cache()

        return jsonify({
            'ok': True,
            'message': f'Файл {safe_name} успешно загружен.',
            'file_info': get_current_file_info(),
        })
    except Exception as e:
        return jsonify({'ok': False, 'error': f'Ошибка обработки файла: {e}'}), 400


@app.route('/api/report', methods=['POST'])
def api_report():
    data = request.json
    try:
        report = build_report(
            data['period_type'],
            data.get('period_value', '1'),
            data['year_new'],
            data['year_old'],
        )
        return jsonify(report)
    except FileNotFoundError as e:
        return jsonify({'error': str(e)}), 400


@app.route('/api/download', methods=['POST'])
def api_download():
    data = request.json
    try:
        report = build_report(
            data['period_type'],
            data.get('period_value', '1'),
            data['year_new'],
            data['year_old'],
        )
        buf = generate_xlsx(report)
        fname = f"comparison_{report['period_label']}_{report['year_new']}_vs_{report['year_old']}.xlsx"
        return send_file(buf, download_name=fname, as_attachment=True,
                         mimetype='application/vnd.openxmlformats-officedocument.spreadsheetml.sheet')
    except FileNotFoundError as e:
        return jsonify({'error': str(e)}), 400


if __name__ == '__main__':
    if XLSX_IN.exists():
        get_tx()   # preload transactions
        get_dds()  # preload DDS sheet
    else:
        print(f'Файл не найден при старте: {XLSX_IN}')
    app.run(debug=False, port=5000)
