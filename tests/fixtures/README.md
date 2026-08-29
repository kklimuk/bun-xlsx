# XLSX compatibility fixtures

This corpus intentionally mixes workbooks IronCalc is expected to support with workbooks that expose
unsupported, malformed, or producer-specific behavior. It must not become an IronCalc-only happy-path
suite.

## IronCalc upstream fixtures (01–20)

Files in `workbooks/` come from IronCalc commit
`195166fef3503eaf56b04767f11e58f5abdc0cb7`, which is dual MIT/Apache-2.0 licensed.
The original paths and reasons for inclusion are:

| Fixture | Original path | Coverage |
|---|---|---|
| 01 | `xlsx/tests/basic_text.xlsx` | Basic types and text |
| 02 | `xlsx/tests/example.xlsx` | Many sheets, chart sheet, hidden sheet, merges, formulas |
| 03 | `xlsx/tests/openpyxl_example.xlsx` | OpenPyXL-produced formula workbook |
| 04 | `xlsx/tests/libreoffice_888_example.xlsx` | LibreOffice-produced financial sheet |
| 05 | `xlsx/tests/missing_r_on_row.xlsx` | Worksheet row missing its usual `r` attribute |
| 06 | `xlsx/tests/optional_xf_id.xlsx` | Optional style identifier |
| 07 | `xlsx/tests/shared_formula_volatile.xlsx` | Shared and volatile formulas |
| 08 | `xlsx/tests/link_test.xlsx` | Hyperlinks and linked formulas |
| 09 | `xlsx/tests/calc_test_no_export/tables.xlsx` | Excel tables and structured references |
| 10 | `xlsx/tests/templates/invoice.xlsx` | Realistic invoice template |
| 11 | `xlsx/tests/templates/gantt_project_tracker.xlsx` | Larger project tracker and date formulas |
| 12 | `xlsx/tests/templates/wordle.xlsx` | Large formula-heavy interactive template |
| 13 | `xlsx/tests/templates/yearly_calendar.xlsx` | Calendar/date layout |
| 14 | `xlsx/tests/dynamic_arrays.xlsx` | Dynamic arrays and spill behavior |
| 15 | `xlsx/tests/calc_tests/defined_names.xlsx` | Workbook- and sheet-scoped names |
| 16 | `xlsx/tests/calc_tests/LOOKUP_AND_REFERENCE/XLOOKUP.xlsx` | Modern lookups and error cases |
| 17 | `xlsx/tests/calc_tests/DATE_AND_TIME/NETWORKDAYS_NETWORKDAYS.INTL.xlsx` | Date calculations |
| 18 | `xlsx/tests/calc_tests/TEXT/UNICODE.xlsx` | Unicode functions and characters |
| 19 | `xlsx/tests/bad_merge_cells.xlsx` | Deliberately malformed merge metadata |
| 20 | `xlsx/tests/conditional_formatting/cf_tests.xlsx` | Conditional-formatting payloads |

The source repository's MIT license is reproduced in `IRONCALC-LICENSE-MIT.txt`.

## Generated independent fixtures (21–28)

Files in `generated/` are produced by `scratches/generate_openpyxl.py` using OpenPyXL 3.1.5. They
cover CSV escaping, uncached formulas, unsupported functions, the 1904 date system, hidden sheets,
sparse wide ranges, typed errors/blanks, and a combination of tables, merges, comments, and
validation.

Regenerate them with:

```sh
python3 scratches/generate_openpyxl.py
```

These generated files are intentionally independent of IronCalc. Some are expected to expose
IronCalc limitations; see `FINDINGS.md` and `tests/corpus.test.ts`.

## Published financial-modeling fixtures (29–30)

These files come from Packt's MIT-licensed
`Hands-On-Financial-Modeling-with-Excel-for-Microsoft-365` repository at commit
`a718ae40267c5c6e42debcea9ca1a6f5f6afffe9`.

| Fixture | Original file | Coverage |
|---|---|---|
| 29 | `Wazobia Global Ltd - ADV FM Final.xlsx` | Four-sheet financial forecast and valuation model with 778 formulas |
| 30 | `WORKINGS FOR FM BOOK.xlsx` | Thirteen sheets, 346 formulas, and 33 formulas with real cross-sheet dependencies |

The source repository's MIT license is reproduced in
`PACKT-FINANCIAL-MODELING-LICENSE-MIT.txt`.
