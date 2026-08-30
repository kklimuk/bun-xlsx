import { describe, expect, test } from "bun:test";
import { readdirSync } from "node:fs";
import { basename, join } from "node:path";
import { WorkbookReader } from "../src/lib.js";

const FIXTURE_ROOT = join(import.meta.dir, "fixtures");
const fixturePaths = ["workbooks", "generated"]
	.flatMap((directory) =>
		readdirSync(join(FIXTURE_ROOT, directory))
			.filter((name) => name.endsWith(".xlsx"))
			.map((name) => join(FIXTURE_ROOT, directory, name)),
	)
	.sort();

const IMPORT_FAILURES = new Map([
	["19-bad-merge-cells.xlsx", 'XML Error: Missing "ref" XML attribute'],
]);

function sheetAt(reader: WorkbookReader, index: number) {
	const sheet = reader.sheets[index];
	if (!sheet) throw new Error(`Fixture is missing worksheet index ${index}`);
	return sheet;
}

function sheetNamed(reader: WorkbookReader, name: string) {
	const sheet = reader.sheets.find((candidate) => candidate.name === name);
	if (!sheet) throw new Error(`Fixture is missing worksheet ${name}`);
	return sheet;
}

describe("XLSX compatibility corpus", () => {
	test("contains at least 20 diverse workbooks", () => {
		expect(fixturePaths.length).toBeGreaterThanOrEqual(20);
		expect(fixturePaths.length).toBe(30);
	});

	for (const path of fixturePaths) {
		const name = basename(path);
		test(name, () => {
			const expectedImportFailure = IMPORT_FAILURES.get(name);
			if (expectedImportFailure) {
				expect(() => new WorkbookReader(path, "computed")).toThrow(
					expectedImportFailure,
				);
				expect(() => new WorkbookReader(path, "formulas")).toThrow(
					expectedImportFailure,
				);
				return;
			}

			const computed = new WorkbookReader(path, "computed");
			const raw = new WorkbookReader(path, "formulas");
			expect(computed.sheets.length).toBeGreaterThan(0);
			expect(raw.sheets.map((sheet) => sheet.name)).toEqual(
				computed.sheets.map((sheet) => sheet.name),
			);

			if (name === "26-sparse-wide-range.xlsx") {
				expect(() =>
					computed.toCsv(sheetAt(computed, 0), 100_000, "none"),
				).toThrow("exceeding --max-cells 100000");
				expect(() => raw.toCsv(sheetAt(raw, 0), 100_000, "none")).toThrow(
					"exceeding --max-cells 100000",
				);
				return;
			}

			for (const sheet of computed.sheets) {
				const csv = computed.toCsv(sheet, 100_000, "none");
				expect(csv === "" || csv.endsWith("\n")).toBe(true);
			}
			for (const sheet of raw.sheets) {
				const csv = raw.toCsv(sheet, 100_000, "none");
				expect(csv === "" || csv.endsWith("\n")).toBe(true);
			}
		});
	}
});

describe("reviewed corpus outputs", () => {
	test("coordinate labels preserve offset worksheet rows and columns", () => {
		const reader = new WorkbookReader(
			join(FIXTURE_ROOT, "workbooks/07-shared-volatile-formula.xlsx"),
			"all",
		);
		const csv = reader.toCsv(sheetAt(reader, 0), 100_000);
		expect(csv.startsWith("Row,C,D\n3,")).toBe(true);
	});

	test("known Excel compatibility differences remain explicit", () => {
		const openpyxl = new WorkbookReader(
			join(FIXTURE_ROOT, "workbooks/03-openpyxl-upstream.xlsx"),
			"computed",
		);
		expect(openpyxl.toCsv(sheetAt(openpyxl, 0), 100_000, "none")).toBe(
			'"Hello, World!",It is what it is\n2,\n',
		);

		const hyperlinks = new WorkbookReader(
			join(FIXTURE_ROOT, "workbooks/08-hyperlinks.xlsx"),
			"computed",
		);
		const hyperlinkCsv = hyperlinks.toCsv(
			sheetNamed(hyperlinks, "Sheet1"),
			100_000,
			"none",
		);
		expect(hyperlinkCsv.match(/#NAME\?/g)).toHaveLength(5);
		expect(hyperlinkCsv).toContain(
			"20. Formula link - external workbook,#ERROR!",
		);

		const networkdays = new WorkbookReader(
			join(FIXTURE_ROOT, "workbooks/17-networkdays.xlsx"),
			"computed",
		);
		const networkdaysCsv = networkdays.toCsv(
			sheetAt(networkdays, 0),
			100_000,
			"none",
		);
		expect(networkdaysCsv).toContain(
			"TRUE,FALSE,#NUM!,#NUM!,Number is boolean",
		);
		expect(networkdaysCsv).toContain("TRUE,03-01-00,-44,-44,Number is boolean");
		expect(networkdaysCsv).toContain("03-01-00,TRUE,44,44,Number is boolean");
	});

	test("CSV escaping preserves commas, quotes, newlines, and Unicode", () => {
		const reader = new WorkbookReader(
			join(FIXTURE_ROOT, "generated/21-text-csv-escaping.xlsx"),
			"computed",
		);
		expect(reader.toCsv(sheetAt(reader, 0), 100_000, "none")).toBe(
			'kind,value\ncomma,"alpha,beta"\nquote,"He said ""hello"""\nnewline,"first line\nsecond line"\nunicode,naïve café — 東京 — 😀\nformula-looking text,=not a formula\n',
		);
	});

	test("computed mode evaluates formula chains while formulas preserves formulas", () => {
		const path = join(FIXTURE_ROOT, "generated/22-formulas-without-cache.xlsx");
		const computed = new WorkbookReader(path, "computed");
		const raw = new WorkbookReader(path, "formulas");
		const computedCsv = computed.toCsv(sheetAt(computed, 0), 100_000, "none");
		const rawCsv = raw.toCsv(sheetAt(raw, 0), 100_000, "none");
		expect(computedCsv).toBe(
			"input,double,running total,,\n3,6,3,,87\n5,10,8,,\n8,16,16,,\n13,26,29,,\n",
		);
		expect(rawCsv).toBe(
			"input,double,running total,,\n3,=A2*2,=SUM($A$2:A2),,=SUM(B2:B5)+C5\n5,=A3*2,=SUM($A$2:A3),,\n8,=A4*2,=SUM($A$2:A4),,\n13,=A5*2,=SUM($A$2:A5),,\n",
		);
	});

	test("unsupported formulas degrade to cell errors without failing the read", () => {
		const reader = new WorkbookReader(
			join(FIXTURE_ROOT, "generated/23-unsupported-functions.xlsx"),
			"computed",
		);
		const csv = reader.toCsv(sheetAt(reader, 0), 100_000, "none");
		expect(csv).toContain("web service,#NAME?");
		expect(csv).toContain("cube,#NAME?");
		expect(csv).toContain("lambda helper,6");
	});

	test("hidden and very-hidden worksheets remain visible in extraction", () => {
		const reader = new WorkbookReader(
			join(FIXTURE_ROOT, "generated/25-hidden-and-very-hidden.xlsx"),
			"computed",
		);
		expect(reader.sheets.map((sheet) => sheet.name)).toEqual([
			"Visible",
			"Hidden",
			"Very Hidden",
		]);
		expect(reader.toCsv(sheetAt(reader, 1), 100_000, "none")).toBe(
			"hidden value\n",
		);
		expect(reader.toCsv(sheetAt(reader, 2), 100_000, "none")).toBe(
			"very hidden value\n",
		);
	});

	test("1904 dates are shifted without changing pure times or durations", () => {
		const reader = new WorkbookReader(
			join(FIXTURE_ROOT, "generated/24-date-system-1904.xlsx"),
			"computed",
		);
		const csv = reader.toCsv(sheetAt(reader, 0), 100_000, "none");
		expect(csv).toBe(
			"date,datetime,date-formatted time,date-formatted duration,time only,elapsed duration\n1904-01-02 00:00:00,2024-02-29 15:45:30,1904-01-01 23:59:58,1904-01-03 01:02:00,12:30:15,49:02\n",
		);
	});

	test("absolute OpenXML relationship targets import through normalization", () => {
		const reader = new WorkbookReader(
			join(FIXTURE_ROOT, "generated/28-table-merge-comment-validation.xlsx"),
			"computed",
		);
		expect(reader.sheets.map((sheet) => sheet.name)).toEqual([
			"Structured data",
		]);
		const csv = reader.toCsv(sheetAt(reader, 0), 100_000, "none");
		expect(csv).toContain("Pen,2,1.5,3");
		expect(csv).toContain("Book,3,8.25,24.75");
		expect(csv).toContain("Bag,1,42,42");
	});

	test("advanced financial model produces coherent forecast values", () => {
		const reader = new WorkbookReader(
			join(FIXTURE_ROOT, "workbooks/29-advanced-financial-model.xlsx"),
			"computed",
		);
		expect(reader.sheets.map((sheet) => sheet.name)).toEqual([
			"Cover",
			"Financial Model",
			"Sheet1",
			"Valuation",
		]);
		expect(sheetNamed(reader, "Cover").range).toBe("A1:F4");
		const csv = reader.toCsv(
			sheetNamed(reader, "Financial Model"),
			100_000,
			"none",
		);
		expect(csv).toContain('GROSS PROFIT,,,"83,028","87,538","65,957"');
		expect(csv).toContain('PROFIT AFTER TAX,,,"18,535","13,787","1,850"');
		expect(csv).not.toMatch(/#(?:NAME\?|VALUE!|REF!|N\/A|DIV\/0!)/);
	});

	test("cross-sheet lookup dependencies compute and remain visible in formulas mode", () => {
		const path = join(
			FIXTURE_ROOT,
			"workbooks/30-cross-sheet-dependencies.xlsx",
		);
		const computed = new WorkbookReader(path, "computed");
		const raw = new WorkbookReader(path, "formulas");
		expect(computed.sheets).toHaveLength(13);
		const computedCsv = computed.toCsv(
			sheetNamed(computed, "SalesReport1"),
			100_000,
			"none",
		);
		const rawCsv = raw.toCsv(sheetNamed(raw, "SalesReport1"), 100_000, "none");
		expect(computedCsv).toContain(
			'Desktop PC,BN001,Mobola,30,"  78,000 ","  2,340,000 ","  65,000 "',
		);
		expect(computedCsv).toContain(
			'Desk Fan,PVC03,Iyabo,36,"  19,200 ","  691,200 ","  16,000 "',
		);
		expect(rawCsv).toContain("INDEX('Prod Dbase'!$D$5:$D$13");
		expect(computedCsv).not.toMatch(/#(?:NAME\?|VALUE!|REF!|N\/A|DIV\/0!)/);
	});
});
