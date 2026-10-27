# Bank-format business plans (UzTMK, 78 mln USD facility)

Generates the three yearly bank-format business plans (`../UzTMK_BankBP_*.xlsx`) from the UzTMK FS models.

```
python3 extract_full.py <FS model.xlsx> full_<key>.json "<model file name>"
python3 -c "from make import make; make('<key>', <price factor>, '../UzTMK_BankBP_<name>.xlsx', final=True, start=<first year>)"
```

Price factors used: sarikul 0.849 (from 2027), tungsten 0.978 (from 2027), moly 0.773 (from 2028).
`adjust.py` applies the factor to sale prices and updates every revenue-linked item; `make.py` builds,
recalculates and writes the source numbers straight into the bank sheets: every input appears once as a value,
everything else is a plain formula link between the bank sheets (the build stops if a formula would mix a source
number with a link).
