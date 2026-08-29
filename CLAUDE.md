# xlsx

Small Bun CLI for deterministic XLSX-to-CSV extraction for AI agents.

## Runtime and structure

- Use Bun, not Node, for the executable and development commands.
- Use Bun's `node:util` `parseArgs` for CLI option parsing. Keep CLI parsing in `src/cli.ts` and workbook logic in `src/lib.ts`.
- The public grammar is `xlsx <read|list|render> FILE [options]`, matching the command-first `docx-cli` family. Bare `xlsx` shows top-level help; there is no implicit read command.
- Every command has dedicated `--help`; top-level help is for command discovery, while command help owns detailed options, examples, and output semantics.
- Query output goes to stdout. Errors go to stderr and return nonzero.
- `read` streams all worksheets in workbook order. Multi-sheet output prefixes each block with `Sheet: ` followed by the CSV-escaped worksheet name and separates blocks by exactly two blank lines. Single-sheet output has no added header. It must remain redirectable with `>`.
- `render` produces PNG only. Microsoft Excel exports each worksheet with a landscape page setup from a staged copy on macOS. Headless LibreOffice is the fallback and preserves saved print orientation; do not use LibreOfficePython or UNO helpers. Bundled PDFium creates page images. The source workbook is unchanged, and image paths go to stdout.
- Releases use one version in `package.json` and a matching immutable `v<version>` tag. The tag publishes checksummed standalone binaries on GitHub and the `@sageling/xlsx` package on npm. `xlsx upgrade` may replace only a standalone binary; package-manager installs must remain package-manager-owned. The embedded installer and release asset must fetch binaries and `SHA256SUMS` from a release, never execute a remotely fetched script.

## Workbook invariants

- IronCalc owns cell input, formula evaluation, and formatted computed values. Do not add a second workbook value engine.
- Computed values are the default. `--values raw` returns literal inputs and formulas.
- Preserve the rectangular content-bearing worksheet range, including empty interior cells. Ignore formatting-only cells outside it; use the declared dimension only as a fallback when content exists but cell references are unavailable.
- Never silently truncate. Refuse sheets above `--max-cells` and archives above the decompression safety limit.
- Keep dependency versions pinned. IronCalc is pre-1.0, so upgrades require raw/computed regression tests.
- IronCalc compatibility normalization operates only on a private temporary import copy. Never mutate the source workbook. Internal absolute relationships may be made relative; 1904 date serials may be converted only for non-formula cells with date-bearing styles. Pure numbers, time-only formats, and elapsed-duration formats must not shift.

## Verification

Run all three before handing off changes:

```sh
bun test
bun run check             # Biome + Knip + TypeScript
bun run build
```

Also exercise the compiled `dist/xlsx` against a multi-sheet workbook in both default computed mode and `--values raw` mode, and smoke-test `render` when LibreOffice is available.
