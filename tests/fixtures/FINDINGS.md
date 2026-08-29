# Fixture output review

Reviewed with IronCalc Node binding 0.8.3 and this CLI's default 100,000-cell-per-sheet limit.

## Summary

- 30 workbooks total; 29 import successfully.
- All 29 importable workbooks support `list`.
- 28 produce both computed and raw CSV under the default output bound.
- Fixture 26 imports and lists correctly, then both reads refuse its 1,703,936-cell `A1:XFD104`
  rectangle. This is the intended bounded-output behavior.
- Fixture 19 fails during IronCalc import, before CSV extraction.

## Outputs that make sense

- Fixture 21 correctly applies RFC-style CSV quoting to commas, quotes, and embedded newlines, while
  preserving Unicode and formula-looking text.
- Fixture 22 has no cached formula results. Computed output evaluates the chain (`3 → 6`, running
  totals through `29`, final result `87`); raw output preserves formulas such as `=A2*2`.
- Fixture 23 degrades unsupported `WEBSERVICE`, `CUBEVALUE`, and `IMAGE` formulas to `#NAME?` without
  aborting the workbook. IronCalc does support and compute the `REDUCE`/`LAMBDA` case as `6`.
- Fixture 25 includes visible, hidden, and very-hidden worksheets in workbook order. This is desirable
  for complete extraction rather than UI-faithful visibility filtering.
- Basic Excel, OpenPyXL, and LibreOffice fixtures all produce coherent output. Formula-heavy upstream
  templates show the expected distinction: raw mode contains formulas; computed mode contains display
  values and number formatting.
- Fixture 29 is a four-sheet financial forecast and valuation model with 778 formulas. Its computed
  P&L produces coherent gross-profit and profit-after-tax series without calculation errors. Its
  formatting-only cells once made the declared Cover range span 645,560 cells; extraction now bounds
  the sheet to its actual content-bearing `A1:F4` range.
- Fixture 30 has 13 worksheets and 33 genuine cross-sheet formulas. Computed `INDEX`/`MATCH` lookups
  agree with Excel's cached values (for example, Desktop PC cost 65,000 and selling price 78,000),
  while raw mode preserves the cross-sheet formulas.

## Compatibility issues fixed by the CLI

### 1904 date system

Fixture 24 uses the valid 1904 workbook date system. IronCalc 0.8.3 otherwise interprets date-styled
serials with the 1900 epoch, producing values exactly 1,462 days early. The CLI normalizes a private
import copy: non-formula numeric cells with date-bearing formats are shifted by 1,462 and the copy's
epoch is set to 1900. It does not shift ordinary numbers, pure time formats, or elapsed durations.
Output now contains the expected 1904-01-02, 2024-02-29 15:45:30, 12:30:15, and 49:02 values.

### Absolute OpenXML relationship targets

Fixture 28 is a valid OpenPyXL workbook whose worksheet relationships use absolute targets such as
`/xl/tables/table1.xml` and `/xl/comments/comment1.xml`. IronCalc 0.8.3 fails to resolve the comment
target directly. The CLI converts internal absolute targets to equivalent source-relative targets in
the private import copy. The workbook now lists and reads successfully, including its computed totals.

The original file is never modified. The normalized copy is created with private permissions in the
OS temporary directory, loaded by IronCalc, and removed immediately.

### Malformed merges fail closed

Fixture 19 deliberately omits the required merge `ref`. IronCalc reports
`XML Error: Missing "ref" XML attribute`; the CLI exits nonzero and emits no partial CSV. This is a
reasonable fail-closed result.

### Worksheet boundaries are labeled

Multi-sheet `read` output prefixes each worksheet block with `Sheet: ` and its CSV-escaped name on a
standalone line, then puts exactly two blank lines between blocks. Hidden sheets are included.
Single-sheet output remains plain CSV without an added name line.
