"""Full extraction of annual data (USD) from a UzTMK FS model for the bank-format business plan.

Every number keeps a reference to its source cell row in the FS model ("sheet!стр.N").
"""
import json
import re
import sys
import warnings

import openpyxl

from extract import Sheet, num

warnings.filterwarnings('ignore')
YRS = list(range(2026, 2038))
ELEMENTS = ['Raw materials', 'Auxiliary materials', 'Electricity', 'Gas / fuel', 'Water', 'Payroll',
            'Social tax', 'Depreciation', 'Repairs', 'Other']
CYR = re.compile('[А-Яа-яЁё]')

# Russian names for model lines that are entered in English only (Sarikul capex register)
RU_FALLBACK = {
    'Tailings storage facility (TSF)': 'Хвостохранилище',
    'Buildings': 'Здания',
    'Tunnel and roads': 'Тоннель и автодороги',
    'Processing plant': 'Обогатительная фабрика',
    'Electricity and water supply': 'Электро- и водоснабжение',
    'Open pit': 'Карьер (открытые горные работы)',
    'Underground mine': 'Подземный рудник',
    'Vehicles': 'Транспортные средства',
    'Other': 'Прочие объекты',
    'Geological and other research to date': 'Геологоразведочные и прочие исследования (выполненные)',
    'Sustaining (maintenance) capex — automatic, % of gross PPE cost':
        'Поддерживающие капвложения (% от первоначальной стоимости ОС)',
}


DEPT_RU = {'Management': 'Руководство', 'Mine Dept': 'Горный отдел', 'Geology Dept': 'Геологический отдел',
           'Processing Dept': 'Обогатительный отдел (фабрика)', 'Maintenance Dept': 'Служба технического обслуживания и ремонта',
           'Assay Lab': 'Пробирно-аналитическая лаборатория', 'HSE Dept': 'Отдел охраны труда, промбезопасности и экологии',
           'Admin Dept — office': 'Административный отдел — офис', 'Admin Dept — site services': 'Административный отдел — обслуживание площадки',
           'HR Dept': 'Отдел кадров', 'Procurement Dept': 'Отдел закупок', 'Finance Dept': 'Финансовый отдел',
           'Legal Counsel Dept': 'Юридический отдел', 'Community Relations Dept': 'Отдел по связям с местным сообществом'}
STAFF_RU = {'expatriates': 'иностранные специалисты', 'Uzbek staff': 'местный персонал'}


def ru_position(text):
    t = str(text).strip()
    for st, st_ru in STAFF_RU.items():
        if t.endswith(' — ' + st):
            dept = t[: -len(' — ' + st)]
            if dept in DEPT_RU:
                return f'{DEPT_RU[dept]} — {st_ru}'
    return ru(t)


def ru(text):
    """Return the Russian part of an 'EN | RU' or 'EN / RU' label (or a fallback translation)."""
    if text is None:
        return ''
    t = str(text).replace('\n', ' / ').strip()
    for sep in (' | ', ' / '):
        if sep in t:
            parts = [p.strip() for p in t.split(sep)]
            cyr = [p for p in parts if CYR.search(p)]
            if cyr:
                return cyr[-1] if sep == ' | ' else sep.join(cyr)
    if CYR.search(t):
        return t
    return RU_FALLBACK.get(t, t)


def ser(sh, ri):
    r = sh.rows[ri]
    return {str(y): (num(r[c]) or 0.0) if c < len(r) else 0.0 for y, c in sh.fy.items()}


def extract(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    S = {ws.title: Sheet(ws) for ws in wb.worksheets
         if ws.title not in ('Outputs>>', 'Inputs>>', 'Calcs>>', 'Lists', 'Glossary')}
    D = {'series': {}, 'src': {}, 'meta': {}, 'products': [], 'varlines': [], 'fixlines': [],
         'positions': [], 'capex': [], 'opexlines': []}

    def put(key, sheet, pattern, required=True, **kw):
        sh = S[sheet]
        hits = sh.findall(pattern)
        # several rows can share a label prefix (e.g. a memo header with zeros): take the first non-empty one
        nz = [h for h in hits if any(abs(v) > 0 for v in ser(sh, h).values())]
        ri = (nz or hits or [None])[0]
        if ri is None:
            if required:
                raise KeyError(f'{sheet}: {pattern}')
            D['series'][key] = {str(y): 0.0 for y in YRS}
            D['src'][key] = f'{sheet}: нет строки'
            return None
        D['series'][key] = ser(sh, ri)
        D['src'][key] = f'{sheet}!стр.{ri + 1}'
        return ri

    # ---------------- F1 general ----------------
    f1 = S['F1 General']
    def f1val(code):
        ri = f1.find(r'^' + re.escape(code) + r' \|')
        return f1.rows[ri][3] if ri is not None else None
    D['meta']['name_en'] = str(f1val('1.01')).split(' | ')[0]
    D['meta']['name_ru'] = ru(f1val('1.01'))
    D['meta']['company'] = ru(f1val('1.03'))
    D['meta']['site'] = ru(f1val('1.04'))
    D['meta']['stage'] = str(f1val('3.01'))
    D['meta']['constr_start'] = str(f1val('5.03'))[:10]
    D['meta']['first_prod'] = str(f1val('5.04'))[:10]
    D['meta']['shareholder'] = str(f1val('4.01'))

    # ---------------- P&L / CF / BS (statements) ----------------
    P, C, B = 'P&L', 'CF', 'BS'
    for k, pat in [('rev', r'^Revenue \| Выручка'), ('cos_cash', r'^Cost of sales — cash production'),
                   ('cos_da', r'^Cost of sales — depreciation'), ('cos', r'^Cost of sales \| Себестоимость продаж'),
                   ('sell', r'^Selling and distribution'), ('ga', r'^General and administrative'),
                   ('ooi', r'^Other operating income'), ('ooe', r'^Other operating expenses'),
                   ('ebit', r'^OPERATING PROFIT \(EBIT\)'), ('fin', r'^Finance costs'),
                   ('pbt', r'^PROFIT BEFORE TAX'), ('tax', r'^Income tax expense'),
                   ('np', r'^PROFIT / \(LOSS\) FOR THE PERIOD'), ('da', r'^Depreciation and amortisation \(total\)'),
                   ('ebitda', r'^EBITDA \| USD'), ('fxrate', r'Average exchange rate'),
                   ('usd_prod', r'USD-denominated production costs'),
                   ('rev_usd', r'Revenue from USD-priced products')]:
        put(k, P, pat)
    for k, pat in [('cgo', r'^Cash generated from operations /'), ('taxpaid', r'^Income tax paid'),
                   ('capex_cf', r'^Purchase of property'), ('d_inv', r'^\(Increase\) / decrease in inventories'),
                   ('d_rec', r'^\(Increase\) / decrease in trade receivables'),
                   ('d_vat', r'^\(Increase\) / decrease in VAT'),
                   ('d_pay', r'^Increase / \(decrease\) in trade and other payables'),
                   ('d_wag', r'^Increase / \(decrease\) in wages')]:
        put(k, C, pat)
    for k, pat in [('bs_inv', r'^Inventories \| Запасы'), ('bs_rec', r'^Trade receivables'),
                   ('bs_vat', r'^VAT recoverable'), ('bs_pay', r'^Trade payables'),
                   ('bs_taxp', r'^Taxes payable'), ('bs_wag', r'^Wages and payroll taxes payable')]:
        put(k, B, pat)

    # ---------------- F12 financing (model's own loan — replaced in the bank plan) ----------------
    F12 = 'F12 Financing'
    put('capbc_dep', F12, r'^Depreciation of capitalised borrowing costs')
    put('capbc', F12, r'^Capitalised into construction in progress')
    put('capex', F12, r'^Capex incl\. contingency \(F8\)')
    put('fx_avg', F12, r'^FX3 \|')
    for key, code in [('wacc', '3b.01'), ('capbc_life', '2b.02')]:
        ri = S[F12].find(r'^' + re.escape(code))
        vals = [num(x) for x in S[F12].rows[ri] if num(x) is not None]
        D['meta'][key] = vals[-1]
        D['src'][key] = f'{F12}!стр.{ri + 1}'
    ri = S[F12].find(r'^N1 \|')
    r = S[F12].rows[ri]
    D['meta']['model_loan'] = [num(x) for x in r if num(x) is not None][:2]

    # ---------------- F10 taxes ----------------
    T = 'F10 Taxes'
    put('cit_rate', T, r'^1\.01 \| Corporate income tax rate')
    put('prop_tax', T, r'^2\.04 \| Property tax')
    put('land_tax', T, r'^2\.07 \| Land tax')
    put('rent_tax', T, r'^6\.24 \| SPECIAL RENT TAX', required=False)
    put('vat_rate', T, r'^4\.01 \| VAT rate')
    put('vat_out', T, r'^4\.08 \| Output VAT')
    put('vat_in_op', T, r'^4\.09 \| Input VAT — operating')
    put('vat_in_capex', T, r'^4\.10 \| Input VAT — capex')
    put('vat_net', T, r'^4\.12 \| NET VAT')
    ri = S[T].find(r'^6\.02 \| Special rent tax rate')
    D['meta']['rent_rate'] = [num(x) for x in S[T].rows[ri] if num(x) is not None][-1] if ri is not None else 0
    ri = S[T].find(r'^2\.02 \| Property tax rate')
    D['meta']['prop_rate'] = [num(x) for x in S[T].rows[ri] if num(x) is not None][-1]

    # ---------------- Assumptions: WC days, tax settings ----------------
    A = S['Assumptions']
    def aval(code):
        for ri, r in enumerate(A.rows):
            if r[1] == code:
                vals = [num(x) for x in r[3:8] if num(x) is not None]
                return (vals[0] if vals else None), ri + 1
        return None, None
    for key, code in [('wip_days', '7.01'), ('fg_days', '7.02'), ('dso', '7.03'), ('dpo', '7.04'),
                      ('wage_lag', '6.01'), ('vat_refund_lag', '6.06'), ('capex_vat_lag', '6.09'),
                      ('ke', '5.02'), ('soc_rate', '3.04'), ('loss_cf_years', '3.07'), ('prop_rate_a', '3.03')]:
        v, rr = aval(code)
        D['meta'][key] = v
        D['src'][key] = f'Assumptions!стр.{rr}'
    f11 = S['F11 Working capital']
    ri = f11.find(r'days held')
    vals = [num(x) for x in f11.rows[ri][5:45] if num(x) is not None] if ri is not None else []
    D['meta']['rm_days'] = max(vals) if vals else None
    D['src']['rm_days'] = f'F11 Working capital!стр.{ri + 1}' if ri is not None else ''

    # ---------------- Dashboard KPIs (model, 12-year horizon, model financing) ----------------
    db = S['Dashboard']
    for key, pat, col in [('npv_model', r'^Project NPV', 3), ('irr_model', r'^Project IRR', 3),
                          ('payback_model', r'^Payback', 3), ('dpayback_model', r'^Discounted payback', 3),
                          ('pi_model', r'^Profitability index', 3), ('tpc_model', r'^Project NPV', 8)]:
        ri = db.find(pat)
        D['meta'][key] = num(db.rows[ri][col])
        D['src'][key] = f'Dashboard!стр.{ri + 1}'
    ri = db.find(r'^Project NPV')
    D['meta']['be_price'] = num(db.rows[ri][14])

    # ---------------- F2/F3 products ----------------
    f2, f3 = S['F2 Production'], S['F3 Prices & Revenue']
    for ri, r in enumerate(f2.rows):
        if ri < 12 or not (isinstance(r[1], str) and re.match(r'^M\d', r[1])):
            continue
        if f2.find(r'NAMEPLATE') is not None and ri > f2.find(r'NAMEPLATE'):
            break
        code = r[1]
        prod = {'code': code, 'name_ru': ru(r[2]), 'unit': r[3] if isinstance(r[3], str) else 't',
                'prod': ser(f2, ri), 'src_prod': f'F2 Production!стр.{ri + 1}'}
        D['products'].append(prod)
    cap0 = f2.find(r'NAMEPLATE')
    for p in D['products']:
        ri = f2.find(r'^' + re.escape(p['code']) + r' \|', start=cap0)
        r = f2.rows[ri]
        p['capacity'] = [num(x) for x in r[3:5] if num(x) is not None][0]
        p['util'] = ser(f2, ri)
        p['src_util'] = f'F2 Production!стр.{ri + 1}'
        s3 = f3.find(r'^2\.\s+SALES VOLUME')
        for key, pat in [('sales_dom', r'Sales — domestic'), ('sales_exp', r'Sales — export'),
                         ('close', r'Closing stock'), ('price', r'Price selected, USD'),
                         ('rev', r'Revenue')]:
            rr = None
            for k in range(s3, len(f3.rows)):
                lab = f3.label(k)
                if f3.rows[k][0] == p['code'] or f3.rows[k][1] == p['code'] or lab.startswith(p['code'] + ' |'):
                    if re.search(pat, lab):
                        rr = k
                        break
            p[key] = ser(f3, rr) if rr is not None else {str(y): 0.0 for y in YRS}
            p['src_' + key] = f'F3 Prices & Revenue!стр.{rr + 1}' if rr is not None else 'нет строки'
        p['market'] = 'export' if sum(p['sales_exp'].values()) > 0 else 'domestic'

    # ---------------- F4 COGS ----------------
    f4 = S['F4 COGS']
    end1 = f4.find(r'Total variable unit cost')
    qhdr = f4.find(r'Input \(cost line\)')
    for ri in range(10, end1):
        r = f4.rows[ri]
        nxt = f4.rows[ri + 1] if ri + 1 < len(f4.rows) else []
        if (isinstance(r[1], str) and re.match(r'^M\d', r[1]) and isinstance(r[3], str)
                and not r[3].strip().startswith('→') and r[4] in ELEMENTS and nxt and nxt[0] in ('USD', 'UZS')):
            uc = f4.rows[ri + 2]
            if not (isinstance(uc[3], str) and '→ unit cost' in uc[3]):
                continue
            line = {'code': r[1], 'name_ru': ru(r[3]), 'element': r[4], 'ccy': nxt[0],
                    'unit': nxt[3], 'price_unit': nxt[4], 'norm': ser(f4, ri), 'price_ccy': ser(f4, ri + 1),
                    'unit_cost': ser(f4, ri + 2), 'src': f'F4 COGS!стр.{ri + 1}-{ri + 3}'}
            # quantity row in the "Input (cost line)" table
            for k in range(qhdr, len(f4.rows)):
                q = f4.rows[k]
                if q[1] == r[1] and q[3] == r[3]:
                    line['qty'] = ser(f4, k)
                    line['src_qty'] = f'F4 COGS!стр.{k + 1}'
                    break
            prod = next(p for p in D['products'] if p['code'] == r[1])
            line['cost'] = {y: line['unit_cost'][y] * prod['prod'][y] for y in line['unit_cost']}
            D['varlines'].append(line)
    for ri, r in enumerate(f4.rows):
        if isinstance(r[1], str) and re.match(r'^3\.\d\d$', r[1]) and isinstance(r[2], str) and r[4] in ELEMENTS:
            D['fixlines'].append({'code': r[1], 'name_ru': ru(r[2]), 'element': r[4], 'ccy': r[0],
                                  'amount_ccy': ser(f4, ri), 'src': f'F4 COGS!стр.{ri + 1}'})
    el0 = f4.find(r'^Cost element')
    for el in ELEMENTS:
        k = next(k for k in range(el0, len(f4.rows)) if f4.rows[k][1] == el and f4.rows[k][4] == 'USD')
        key = 'el_' + el.split()[0].lower().replace('/', '')
        D['series'][key] = ser(f4, k)
        D['src'][key] = f'F4 COGS!стр.{k + 1}'
    put('el_total', 'F4 COGS', r'^TOTAL PRODUCTION COST by element')
    put('met', 'F4 COGS', r'^Total extraction tax, USD', required=False)
    put('var_cogs', 'F4 COGS', r'^Total variable COGS')
    put('fix_cogs', 'F4 COGS', r'^Total fixed production costs')
    put('cogs_cash', 'F4 COGS', r'^TOTAL COGS \(cash, excl\. D&A\)')

    # ---------------- F5 opex ----------------
    f5 = 'F5 Opex'
    put('opex_cash', f5, r'^TOTAL OPEX \(cash, excl\. D&A\)')
    put('opex_ga_da', f5, r'^D&A — G&A')
    put('pers_ga', f5, r'^Personnel — G&A / corporate positions')
    put('opex_fixed', f5, r'^Total fixed opex')
    put('opex_var', f5, r'^Total variable opex')
    put('cc_sales', f5, r'^Sales & marketing \[|^Sales & marketing$', required=False)
    sh5 = S[f5]
    for ri, r in enumerate(sh5.rows):
        if isinstance(r[1], str) and re.match(r'^[12]\.\d\d$', r[1]) and isinstance(r[2], str) and r[2].strip():
            amt = None
            if r[1].startswith('2.'):
                # variable line: amount row follows the % row
                amt = ser(sh5, ri + 1)
            D['opexlines'].append({'code': r[1], 'name_ru': ru(r[2]), 'ccy': r[0],
                                   'amount': amt if amt else ser(sh5, ri), 'src': f'{f5}!стр.{ri + 1}'})
    # ---------------- F6 contingency ----------------
    put('contingency', 'F6 Other items', r'^Opex contingency \| Резерв')
    put('f6_net', 'F6 Other items', r'^NET P&L IMPACT OF FORM 6')

    # ---------------- F9 personnel ----------------
    f9 = S['F9 Personnel']
    for ri, r in enumerate(f9.rows):
        lab = f9.label(ri)
        if 'Headcount' in lab and isinstance(r[2], str) and r[2].strip():
            cost_ri = next(k for k in range(ri + 1, ri + 4) if 'Salary cost, USD' in f9.label(k))
            D['positions'].append({'name_ru': ru_position(r[2]), 'cc': r[3], 'ccy': f9.rows[ri + 1][0],
                                   'headcount': ser(f9, ri), 'monthly_ccy': ser(f9, ri + 1),
                                   'cost': ser(f9, cost_ri), 'src': f'F9 Personnel!стр.{ri + 1}-{cost_ri + 1}'})

    # ---------------- F8 capex ----------------
    f8 = S['F8 Capex']
    ph0 = f8.find(r'^2\.\s+TANGIBLE CAPEX PHASING')
    dp0 = f8.find(r'^6\.\s+DEPRECIATION')
    for ri in range(20, ph0):
        r = f8.rows[ri]
        if isinstance(r[1], str) and re.match(r'^[TI]\d+$', r[1]) and isinstance(r[2], str) and r[2].strip():
            code = r[1]
            pr = next(k for k in range(ph0, dp0) if f8.rows[k][1] == code)
            dr = next(k for k in range(dp0, len(f8.rows)) if f8.rows[k][0] == code or f8.rows[k][1] == code)
            D['capex'].append({'code': code, 'name_ru': ru(r[2]), 'name_en': str(r[2]).split(' | ')[0],
                               'cls': r[3], 'cc': r[4], 'life': num(r[8]), 'ccy': f8.rows[pr][0],
                               'avail': str(r[11])[:10], 'total': num(r[15]),
                               'phase': ser(f8, pr), 'dep': ser(f8, dr),
                               'src': f'F8 Capex!стр.{ri + 1},{pr + 1},{dr + 1}'})
    put('dep_new_total', 'F8 Capex', r'^TOTAL D&A — new assets')
    return D


if __name__ == '__main__':
    D = extract(sys.argv[1])
    D['meta']['file'] = sys.argv[3] if len(sys.argv) > 3 else sys.argv[1].split('/')[-1]
    json.dump(D, open(sys.argv[2], 'w'), ensure_ascii=False, indent=1)
    s = D['series']
    print(D['meta'])
    print('products', [(p['code'], p['name_ru'], p['capacity'], p['market'], round(p['rev']['2029'])) for p in D['products']])
    print('varlines', len(D['varlines']), 'fixlines', len(D['fixlines']), 'positions', len(D['positions']),
          'capex', len(D['capex']), 'opexlines', len(D['opexlines']))
    for y in ['2027', '2029']:
        vsum = {}
        for l in D['varlines']:
            vsum[l['element']] = vsum.get(l['element'], 0) + l['cost'][y]
        print(y, 'var by element', {k: round(v) for k, v in vsum.items()}, 'var_cogs', round(s['var_cogs'][y]))
        els = sum(s[k][y] for k in s if k.startswith('el_') and k != 'el_total')
        print(y, 'elements sum', round(els), 'el_total', round(s['el_total'][y]))
        pc = sum(p['cost'][y] for p in D['positions'] if p['cc'] != 'G&A')
        pg = sum(p['cost'][y] for p in D['positions'] if p['cc'] == 'G&A')
        print(y, 'pers prod gross', round(pc), 'el_payroll', round(s['el_payroll'][y]), '| pers G&A gross', round(pg),
              'pers_ga(incl tax)', round(s['pers_ga'][y]))
        cx = sum(c['dep'][y] for c in D['capex'])
        print(y, 'capex dep lines', round(cx), 'dep_new_total', round(s['dep_new_total'][y]),
              'P&L da - capbc', round(s['da'][y] - s['capbc_dep'][y]))
        print(y, 'capex lines', round(sum(c['phase'][y] for c in D['capex'])), 'capex', round(s['capex'][y]))
        print(y, 'rev products', round(sum(p['rev'][y] for p in D['products'])), 'rev', round(s['rev'][y]),
              ' price*sales', round(sum(p['price'][y] * (p['sales_dom'][y] + p['sales_exp'][y]) for p in D['products'])))
