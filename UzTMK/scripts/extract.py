"""Extract annual data (USD) from UzTMK FS models into JSON for the bank-format business plan."""
import openpyxl, re, json, sys, warnings
warnings.filterwarnings('ignore')
YEARS = list(range(2026, 2038))

def num(v):
    if v is None or v == '': return None
    if isinstance(v, (int, float)): return float(v)
    try: return float(v)
    except Exception: return None

class Sheet:
    def __init__(self, ws):
        self.rows = [list(r) for r in ws.iter_rows(max_col=60, values_only=True)]
        self.fy = {}
        for r in self.rows[:20]:
            for i, v in enumerate(r):
                if isinstance(v, str):
                    m = re.match(r'FY(20\d\d)', v.strip())
                    if m and int(m.group(1)) not in self.fy: self.fy[int(m.group(1))] = i
        if not self.fy:  # F8-style: header deeper
            for r in self.rows:
                for i, v in enumerate(r):
                    if isinstance(v, str):
                        m = re.match(r'FY(20\d\d)', v.strip())
                        if m and int(m.group(1)) not in self.fy: self.fy[int(m.group(1))] = i
                if self.fy: break
    def label(self, ri, upto=7):
        return ' | '.join(str(x).replace('\n', ' / ') for x in self.rows[ri][:upto] if isinstance(x, str))
    def find(self, pattern, start=0, nth=0, upto=7):
        rx = re.compile(pattern)
        k = 0
        for ri in range(start, len(self.rows)):
            if rx.search(self.label(ri, upto)):
                if k == nth: return ri
                k += 1
        return None
    def findall(self, pattern, start=0, end=None, upto=7):
        rx = re.compile(pattern)
        return [ri for ri in range(start, end or len(self.rows)) if rx.search(self.label(ri, upto))]
    def series(self, ri):
        r = self.rows[ri]
        return {y: (num(r[c]) or 0.0) if c < len(r) else 0.0 for y, c in self.fy.items()}
    def get(self, pattern, **kw):
        ri = self.find(pattern, **kw)
        if ri is None: raise KeyError(pattern)
        return self.series(ri), ri + 1

def extract(path):
    wb = openpyxl.load_workbook(path, read_only=True, data_only=True)
    S = {ws.title: Sheet(ws) for ws in wb.worksheets if ws.title not in ('Outputs>>', 'Inputs>>', 'Calcs>>', 'Lists', 'Glossary')}
    out = {'src': {}, 'series': {}, 'meta': {}}
    def put(key, sheet, pattern, **kw):
        s, r = S[sheet].get(pattern, **kw)
        out['series'][key] = {str(y): v for y, v in s.items()}
        out['src'][key] = f"{sheet}!стр.{r}"
    # P&L (output)
    P = 'P&L'
    put('rev', P, r'^Revenue \| Выручка')
    put('cos_cash', P, r'^Cost of sales — cash production')
    put('cos_da', P, r'^Cost of sales — depreciation')
    put('cos', P, r'^Cost of sales \| Себестоимость продаж')
    put('sell', P, r'^Selling and distribution')
    put('ga', P, r'^General and administrative')
    put('ooi', P, r'^Other operating income')
    put('ooe', P, r'^Other operating expenses')
    put('ebit', P, r'^OPERATING PROFIT \(EBIT\)')
    put('fin', P, r'^Finance costs')
    put('fx', P, r'^Foreign exchange gain')
    put('pbt', P, r'^PROFIT BEFORE TAX')
    put('tax', P, r'^Income tax expense')
    put('np', P, r'^PROFIT / \(LOSS\) FOR THE PERIOD')
    put('da', P, r'^Depreciation and amortisation \(total\)')
    put('ebitda', P, r'^EBITDA \| USD')
    put('fxrate', P, r'Average exchange rate')
    # CF
    C = 'CF'
    put('cgo', C, r'^Cash generated from operations /')
    put('taxpaid', C, r'^Income tax paid')
    put('cfo', C, r'^NET CASH FROM OPERATING ACTIVITIES')
    put('capex_cf', C, r'^Purchase of property')
    put('fcff', C, r'^FREE CASH FLOW TO THE FIRM')
    put('d_inv', C, r'^\(Increase\) / decrease in inventories')
    put('d_rec', C, r'^\(Increase\) / decrease in trade receivables')
    put('d_vat', C, r'^\(Increase\) / decrease in VAT')
    put('d_pay', C, r'^Increase / \(decrease\) in trade and other payables')
    put('d_wag', C, r'^Increase / \(decrease\) in wages')
    # other CGO items (if any) captured as residual later
    # BS
    B = 'BS'
    put('bs_inv', B, r'^Inventories \| Запасы')
    put('bs_rec', B, r'^Trade receivables')
    put('bs_vat', B, r'^VAT recoverable')
    put('bs_pay', B, r'^Trade payables')
    put('bs_taxp', B, r'^Taxes payable')
    put('bs_wag', B, r'^Wages and payroll taxes payable')
    put('bs_ppe', B, r'^Property, plant and equipment — cost')
    # F12
    F12 = 'F12 Financing'
    put('capbc_dep', F12, r'^Depreciation of capitalised borrowing costs')
    put('capbc', F12, r'^Capitalised into construction in progress')
    put('capex', F12, r'^Capex incl\. contingency \(F8\)')
    ri = S[F12].find(r'^3b\.01'); r = S[F12].rows[ri]
    out['meta']['wacc'] = [num(x) for x in r if num(x) is not None and 0 < num(x) < 1][-1]
    out['src']['wacc'] = f"{F12}!стр.{ri+1}"
    ri = S[F12].find(r'2b\.02'); r = S[F12].rows[ri]
    out['meta']['capbc_life'] = [num(x) for x in r if num(x) is not None][-1]
    out['src']['capbc_life'] = f"{F12}!стр.{ri+1}"
    # F10 CIT rate
    put('cit_rate', 'F10 Taxes', r'^1\.01 \| Corporate income tax rate')
    put('prop_tax', 'F10 Taxes', r'^2\.04 \| Property tax')
    put('land_tax', 'F10 Taxes', r'^2\.07 \| Land tax')
    try: put('rent_tax', 'F10 Taxes', r'^6\.24 \| SPECIAL RENT TAX')
    except KeyError: pass
    put('vat_out', 'F10 Taxes', r'^4\.08 \| Output VAT')
    put('vat_in_op', 'F10 Taxes', r'^4\.09 \| Input VAT — operating')
    put('vat_in_capex', 'F10 Taxes', r'^4\.10 \| Input VAT — capex')
    put('vat_net', 'F10 Taxes', r'^4\.12 \| NET VAT')
    put('vat_rate', 'F10 Taxes', r'^4\.01 \| VAT rate')
    return out, S

if __name__ == '__main__':
    out, S = extract(sys.argv[1])
    json.dump(out, open(sys.argv[2], 'w'), ensure_ascii=False, indent=1)
    for k, v in out['series'].items():
        print(f"{k:14s} {out['src'][k]:28s}", ' '.join(f"{v[str(y)]/1e6:9.3f}" if abs(v[str(y)])>1000 else f"{v[str(y)]:9.4f}" for y in range(2026, 2034)))
    print(out['meta'])
