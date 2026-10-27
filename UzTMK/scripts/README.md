# Bank-format business plans (UzTMK, 78 mln USD facility)

Generates the three yearly bank-format business plans (`../UzTMK_BankBP_*.xlsx`) from the UzTMK FS models.

```
python3 extract_full.py <FS model.xlsx> full_<key>.json "<original model file name>"
python3 build_bank.py full_<key>.json <sarikul|tungsten|moly> ../UzTMK_BankBP_<name>.xlsx
```

- `extract_full.py` reads annual values (USD) from the FS model and keeps the source sheet/row of every number.
- `build_bank.py` writes the bank template structure (Приложения 1–16, ИНФОРМАЦИЯ, NPV) with live formulas;
  the loan split is taken from the 30% EPC-advance file and scaled to 78.0 mln USD.
- Recalculate the output (e.g. LibreOffice) before reading cached values; sheet «Проверки» must show «ВСЕ OK».
