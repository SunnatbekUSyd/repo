"""Build one bank-format business plan with a sale-price factor, recalculate, optionally finalise.

Finalising replaces every reference to the helper sheet «Исх_данные» by its value (formulas that only read
that sheet become plain numbers) and deletes the helper sheets «Исх_данные» and «Проверки».
"""
import json
import re
import subprocess
import sys

import openpyxl

from adjust import adjust
from build_bank import build

RECALC = ('/root/.claude/skills/synced/e9bc6f8f-14be-4161-9009-5f9053247358_f632cb52-ba1d-4512-90e0-a493e985a139'
          '/xlsx/scripts/recalc.py')
SRC = 'Исх_данные'
REF = re.compile(r"'?Исх_данные'?!\$?([A-Z]+)\$?(\d+)")


def recalc(path):
    out = subprocess.run(['python3', RECALC, path, '180'], capture_output=True, text=True, timeout=400).stdout
    res = json.loads(out)
    assert res.get('status') == 'success', res
    return res


def metrics(path):
    wb = openpyxl.load_workbook(path, data_only=True)
    m = {}
    ws = wb['NPV_IRR']
    for row in ws.iter_rows(min_col=2, max_col=4):
        lab = row[0].value
        if isinstance(lab, str):
            if lab.startswith('Дисконтированный срок'):
                m['dpb'] = row[2].value
            elif lab.startswith('NPV'):
                m['npv'] = row[2].value
            elif lab.startswith('IRR'):
                m['irr'] = row[2].value
            elif lab.startswith('Простой срок'):
                m['pb'] = row[2].value
    if 'Проверки' in wb.sheetnames:
        ws = wb['Проверки']
        m['checks'] = {r[1].value: r[4].value for r in ws.iter_rows(min_row=5) if r[1].value}
    return m


def finalise(path):
    wf = openpyxl.load_workbook(path)
    wv = openpyxl.load_workbook(path, data_only=True)
    src = wf[SRC]
    for ws in wf.worksheets:
        if ws.title in (SRC, 'Проверки'):
            continue
        for row in ws.iter_rows():
            for c in row:
                f = c.value
                if not (isinstance(f, str) and f.startswith('=') and SRC in f):
                    continue
                c.value = wv[ws.title][c.coordinate].value      # input from the feasibility study: keep the number
    del wf[SRC]
    del wf['Проверки']
    wf.active = 0
    wf.save(path)
    left = [ws.title for ws in wf.worksheets for row in ws.iter_rows() for c in row
            if isinstance(c.value, str) and ('Исх_данные' in c.value or 'Проверки!' in c.value)]
    assert not left, left[:5]


def make(key, k, out, final=False, model_file=None, start=2027):
    D = json.load(open(f'full_{key}.json'))
    if model_file:
        D['meta']['file'] = model_file
    build(adjust(D, k, start), key, out)
    recalc(out)
    m = metrics(out)
    if final:
        finalise(out)
        r = recalc(out)
        m['formulas'] = r['total_formulas']
    return m


if __name__ == '__main__':
    print(make(sys.argv[1], float(sys.argv[2]), sys.argv[3], final=len(sys.argv) > 4))
