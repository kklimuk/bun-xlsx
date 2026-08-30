export const TOP_HELP = `xlsx — inspect .xlsx workbooks as CSV or rendered images

Usage:
  xlsx <command> [options]

Commands:
  read  Write every worksheet as CSV; values with formulas are the default
  list  Write worksheet names, ranges, dimensions, and indexes as JSON
  render  Render workbook or worksheet print pages as PNG images
  upgrade  Upgrade a standalone release binary (no workbook argument)

Run \`xlsx <command> --help\` for command-specific help.
Run \`xlsx upgrade --help\` for standalone upgrade options.
`;

export const READ_HELP = `xlsx read — write worksheet contents as CSV

Usage:
  xlsx read <workbook.xlsx> [options]

Examples:
  xlsx read workbook.xlsx
  xlsx read workbook.xlsx > workbook.csv
  xlsx read workbook.xlsx --sheet Sheet1 > sheet.csv
  xlsx read workbook.xlsx --values formulas
  xlsx read workbook.xlsx --labels none
  xlsx read workbook.xlsx --max-cells 500000

Options:
  --sheet NAME        Read only the exactly named worksheet
  --values MODE       all (default) emits formatted values with formulas as
                      VALUE⟦=FORMULA⟧; computed emits values only; formulas emits
                      literal inputs and formulas. An empty formula result is
                      emitted as ⟦=FORMULA⟧
  --labels MODE       coordinates (default) adds Excel column letters and row
                      numbers; none emits only worksheet cell values
  --max-cells N       Refuse any selected sheet whose rectangular used range is
					  larger than N cells (default: 100000); formatting-only cells
					  outside the content rectangle are ignored
  -h, --help          Show this help

Output:
  CSV is written directly to stdout. By default every worksheet is emitted in
  workbook order. For multi-sheet output, each block starts with \`Sheet: NAME\`
  on a standalone line and blocks have exactly two blank lines between them. A
  single selected sheet has no name header. Redirect stdout with \`>\` to save it.
  Coordinate labels use the worksheet's actual rows and columns. Empty cells
  inside each sheet's used range are preserved.
  Hidden and very-hidden sheets are included when reading all sheets. IronCalc
  does not implement every Excel function; use render for visual cross-checking.
  Output is lossless and may contain formula-leading text; do not open CSV from
  an untrusted workbook in spreadsheet software.
  Errors are written to stderr and return a nonzero exit status.
`;

export const LIST_HELP = `xlsx list — describe workbook worksheets as JSON

Usage:
  xlsx list <workbook.xlsx> [options]

Example:
  xlsx list workbook.xlsx

Options:
  -h, --help          Show this help

Output:
  A JSON object is written to stdout. Its sheets array follows workbook order and
	reports each worksheet's zero-based index, name, content-bearing rectangular range, row count,
  column count, and cell count. No worksheet values are emitted.
`;

export const RENDER_HELP = `xlsx render — render workbook print pages as PNG images

Usage:
  xlsx render <workbook.xlsx> [options]

Examples:
  xlsx render workbook.xlsx
  xlsx render workbook.xlsx --out ./pages --pages 1-3
  xlsx render workbook.xlsx --dpi 200
  xlsx render workbook.xlsx --sheet Sheet1 --out ./sheet1-pages

Options:
  --out DIR           Output directory (default: ./<workbook>-pages, or
                      ./<workbook>-<sheet>-pages with --sheet)
  --sheet NAME        Render only the exactly named worksheet
  --dpi N             Image resolution from 36 to 600 (default: 150)
  --pages N[-M]       Render one page or an inclusive page range; with --sheet,
                      page numbers are relative to that worksheet
  -h, --help          Show this help

Output:
  Microsoft Excel is used when available on macOS and applies landscape page
  setup before exporting workbook print pages. LibreOffice is the fallback and
  preserves the workbook's print orientation. Bundled PDFium writes one PNG per
  page. Image paths are written to stdout. The source workbook is not modified.
`;
