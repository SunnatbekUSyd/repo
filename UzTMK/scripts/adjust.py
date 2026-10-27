"""Apply a sale-price factor k (from 2027) to the extracted FS-model data and update every revenue-linked item.

Revenue-linked items: product prices (rounded to whole USD), revenue, selling expenses (% of revenue),
opex contingency (on the selling-expense change), extraction tax (% of revenue), output VAT, trade receivables,
special rent tax (recomputed with its loss pool), deemed income for income tax, operating cash flow before tax.
Income tax is carried as an adjustment to the tax base: tax = rate x (profit before tax + tax_adj).
"""
import copy

Y = [str(y) for y in range(2026, 2038)]


def adjust(D, k, start=2027):
    D = copy.deepcopy(D)
    s = D['series']
    rev_old = dict(s['rev'])
    # prices and revenue by product
    rev_new = {y: 0.0 for y in Y}
    for p in D['products']:
        for y in Y:
            if int(y) >= start and p['price'][y]:
                p['price'][y] = round(p['price'][y] * k)
            p['rev'][y] = p['price'][y] * (p['sales_dom'][y] + p['sales_exp'][y])
            rev_new[y] += p['rev'][y]
    r = {y: (rev_new[y] / rev_old[y] if rev_old[y] else 1.0) for y in Y}
    d_rev = {y: rev_new[y] - rev_old[y] for y in Y}

    def scale(key):
        old = dict(s[key])
        s[key] = {y: old[y] * r[y] for y in Y}
        return {y: s[key][y] - old[y] for y in Y}

    d_sell = scale('sell') if any(s['sell'][y] for y in Y) else {y: 0.0 for y in Y}   # stored negative
    # contingency is a % of (cash COGS + opex): only the selling-expense change feeds it
    base = {y: s['cogs_cash'][y] + s['opex_cash'][y] for y in Y}
    crate = {y: (-s['contingency'][y] / base[y] if base[y] else 0.0) for y in Y}
    d_cont = {y: crate[y] * d_sell[y] for y in Y}                                       # negative (lower cost)
    s['contingency'] = {y: s['contingency'][y] + d_cont[y] for y in Y}
    d_met = scale('met')
    for key in ('el_other', 'el_total'):
        s[key] = {y: s[key][y] + d_met[y] for y in Y}
    s['cos'] = {y: s['cos'][y] - d_met[y] for y in Y}                                   # cos stored negative
    s['cogs_cash'] = {y: s['cogs_cash'][y] + d_met[y] for y in Y}
    s['opex_var'] = {y: s['opex_var'][y] - d_sell[y] for y in Y}                        # opex_var positive
    s['opex_cash'] = {y: s['opex_cash'][y] - d_sell[y] for y in Y}
    d_vout = scale('vat_out')
    s['vat_net'] = {y: s['vat_net'][y] + d_vout[y] for y in Y}
    scale('rev_usd')
    s['rev'] = rev_new

    # special rent tax: object = revenue - cash COGS - extraction tax - capex; losses indexed yearly
    d_rent = {y: 0.0 for y in Y}
    if any(s['rent_tax'][y] for y in Y):
        rate = D['meta']['rent_rate']
        idx = D['meta']['rent_index']
        pool = s['rent_loss_close']['2027']
        for y in Y:
            if int(y) < 2028:
                continue
            obj = s['rent_obj'][y] + d_rev[y] - d_met[y]
            used = min(pool, max(obj, 0.0))
            new_tax = rate * (max(obj, 0.0) - used)
            pool = (pool - used) * (1 + idx) + max(-obj, 0.0)
            d_rent[y] = new_tax - s['rent_tax'][y]
            s['rent_tax'][y] = new_tax
    s['ooe'] = {y: s['ooe'][y] + d_cont[y] - d_rent[y] for y in Y}                     # ooe stored negative

    # EBITDA / EBIT / PBT of the model financing case, operating cash before tax
    d_ebitda = {y: d_rev[y] + d_sell[y] - d_met[y] + d_cont[y] - d_rent[y] for y in Y}
    for key in ('ebitda', 'ebit'):
        s[key] = {y: s[key][y] + d_ebitda[y] for y in Y}
    old_rec = dict(s['bs_rec'])
    s['bs_rec'] = {y: old_rec[y] * r[y] for y in Y}
    d_rec = {y: s['bs_rec'][y] - old_rec[y] for y in Y}
    s['cgo'] = {y: s['cgo'][y] + d_ebitda[y] - (d_rec[y] - (d_rec[str(int(y) - 1)] if y != '2026' else 0.0)) for y in Y}

    # income tax base adjustment (deemed income etc.): implied by the model, deemed income scales with revenue
    deemed = s.get('deemed', {y: 0.0 for y in Y})
    adj = {}
    for y in Y:
        rate = s['cit_rate'][y]
        implied = (-s['tax'][y] / rate - s['pbt'][y]) if rate else 0.0
        adj[y] = implied + deemed[y] * (r[y] - 1)
        if abs(adj[y]) < 1:
            adj[y] = 0.0
    s['tax_adj'] = adj
    D['src']['tax_adj'] = 'расчёт'
    D['meta']['price_factor'] = k
    D['meta']['price_start'] = start
    return D
