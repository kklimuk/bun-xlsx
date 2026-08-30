# xlsx

A small, deterministic XLSX-to-CSV reader intended for use behind an agent harness's `Read` tool.
It uses IronCalc to import cell contents and recalculate formulas.

## Installation

With Bun installed, install the npm package (the package is scoped, but the command is simply
`xlsx`):

```sh
bun add -g @sageling/xlsx
```

Or install a standalone binary with no Bun runtime. Each GitHub release includes binaries for
Linux x64/arm64, macOS x64/arm64, and Windows x64, plus a `SHA256SUMS` manifest:

```sh
curl -fsSLO https://github.com/kklimuk/bun-xlsx/releases/latest/download/install.sh
curl -fsSLO https://github.com/kklimuk/bun-xlsx/releases/latest/download/SHA256SUMS
# macOS: grep ' install.sh$' SHA256SUMS | shasum -a 256 -c -
# Linux: grep ' install.sh$' SHA256SUMS | sha256sum -c -
sh install.sh
```

The installer defaults to `$HOME/.local/bin`; set `PREFIX` or `VERSION=v0.1.0` to override it.
`xlsx upgrade` updates a standalone binary in place, while npm/Bun installations remain owned by
their package manager and should be updated with `bun add -g @sageling/xlsx` or
`npm install -g @sageling/xlsx`.

```sh
xlsx read workbook.xlsx
xlsx read workbook.xlsx --values formulas
xlsx read workbook.xlsx --sheet Sheet1
xlsx list workbook.xlsx
xlsx render workbook.xlsx
xlsx --version
xlsx upgrade --dry-run
```

`read` uses IronCalc to recalculate supported formulas and writes every worksheet as CSV to stdout.
By default, formula cells contain both the formatted value and formula, such as
`22640⟦=SUM(H15:H24)⟧`. Multi-sheet output has a `Sheet: NAME` line and two blank lines between
worksheet blocks. A single-sheet workbook or a
sheet selected with `--sheet` has no added name header. Hidden and very-hidden sheets are included.
IronCalc does not implement every Excel function, so use `render` when visual cross-checking matters.
Use normal shell redirection when a file is wanted:

```sh
xlsx read workbook.xlsx > workbook.csv
```

`--values computed` emits values only; `--values formulas` emits literal inputs and formulas.
`--values all` explicitly selects the default combined form. `--sheet NAME`
limits the output to one worksheet. CSV covers each worksheet's content-bearing rectangular range,
including empty cells inside that rectangle. Formatting-only cells outside it are ignored, avoiding
massive output from Excel's commonly inflated declared dimensions.

By default, CSV adds a `Row` column, Excel column letters, and actual worksheet row numbers so formula
references are easy to locate. Use `--labels none` for cell-only output. To produce one conventional
CSV table, also select one sheet (unless the workbook contains only one sheet).

An empty formula result is emitted without a value prefix, for example `⟦=Engine!A3⟧`. CSV output is
lossless and can begin with spreadsheet formula characters; do not open redirected output from an
untrusted workbook in Excel or LibreOffice.

The default 100,000-cell-per-sheet limit prevents an accidentally inflated worksheet range from
producing unbounded output. Override it explicitly with `--max-cells` when needed. XLSX archives
larger than 256 MiB compressed or expanded beyond 256 MiB are refused before decompression.

## Rendering

CSV cannot expose formatting, charts, merged-cell layout, print areas, or other visual structure.
`render` exports with Microsoft Excel when it is available on macOS. Excel applies a landscape page
setup to every sheet in a staged copy. LibreOffice is the cross-platform fallback and currently
preserves the workbook's saved print orientation. Bundled PDFium writes a PNG for each page:

```sh
xlsx render workbook.xlsx
xlsx render workbook.xlsx --out ./pages --pages 1-3 --dpi 200
```

The default output directory is `./<workbook>-pages`, and rendered image paths are printed to stdout.
Safety limits reject PDFs over 256 MiB, selections over 200 pages, pages over 50 megapixels, total
renders over 500 megapixels, or PNG output over 512 MiB.
Rendering passes the workbook to an installed Office application; isolate that application when
processing hostile files in a server environment.

## Verification

The default suite is entirely non-GUI and includes corpus extraction, complete golden CSVs, CLI
failures, concurrent reads, ZIP safety, worksheet normalization, PDFium rasterization, page ranges,
and stale-output protection:

```sh
bun run check
bun test
```

Additional non-GUI checks exercise the compiled executable outside the repository and enforce broad
performance budgets:

```sh
bun run test:binary
bun run test:performance
```

The Excel integration test is opt-in because it automates the desktop application. It verifies true
landscape output, source immutability, temporary-file cleanup, and stale-page removal:

```sh
bun run test:excel
```

CI runs the normal suite and build on Linux, macOS, and Windows. Binary relocation and performance
checks run on Linux and macOS. Excel automation remains a manual macOS check. Forced-landscape
LibreOffice export remains an explicit compatibility gap; the fallback does not launch Python or UNO
helpers.
Microsoft Excel or LibreOffice is required only for `render`; `read` and `list` are self-contained.
The source workbook is never modified.

Before IronCalc import, the CLI transparently normalizes two valid XLSX variants in a private temporary
copy: absolute internal OpenXML relationship targets and the 1904 workbook date system. The original
workbook is never changed, and the temporary copy is removed immediately after import.

## Development

```sh
bun install
bun test
bun run check
bun run build
```

`bun run build` produces a local standalone executable at `dist/xlsx`; `bun run build:npm` produces
the bundled JavaScript package entrypoint. The IronCalc Node binding is pre-1.0; keep its version
pinned and backtest formula-heavy workbooks before upgrading it.

## Releasing

The package version and Git tag are one release identity. Update `package.json` and `bun.lock`, merge
to `main`, then push the matching tag (for example `v0.1.0`). The tag triggers both the checksummed
GitHub binary release and npm trusted publishing. Do not reuse or move a published tag.

The GitHub repository is `kklimuk/bun-xlsx`; the npm package is `@sageling/xlsx`. Configure npm's
trusted publisher for repository `kklimuk/bun-xlsx`, workflow `release.yml`, and GitHub environment
`Publishing`. If npm requires the new package to exist before that publisher can be configured,
publish the initial version manually, configure trusted publishing, then bump the version before
creating the first automated release tag.
