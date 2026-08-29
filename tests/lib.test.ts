import { describe, expect, test } from "bun:test";
import { csvField, testing } from "../src/lib.js";

describe("csvField", () => {
	test("leaves simple values alone", () =>
		expect(csvField("hello")).toBe("hello"));
	test("quotes commas, quotes, and newlines", () => {
		expect(csvField('a,"b"\nc')).toBe('"a,""b""\nc"');
	});
});

describe("worksheet structure", () => {
	test("decodes XML entities and worksheet ranges", () => {
		expect(testing.decodeXml("A&amp;B &#x1f600; &#65; &quot;x&quot;")).toBe(
			'A&B 😀 A "x"',
		);
		expect(testing.decodeRange("$B$2:AA10")).toEqual({
			startRow: 1,
			startColumn: 1,
			endRow: 9,
			endColumn: 26,
		});
		expect(() => testing.decodeRange("A0")).toThrow("Invalid worksheet range");
	});

	test("finds only content-bearing cells", () => {
		const xml = [
			'<c r="A1" s="1"></c>',
			'<c r="C3"><v>1</v></c>',
			'<c r="AA10"><f>SUM(A1:A2)</f></c>',
			'<c r="XFD1048576"><is><t>end</t></is></c>',
		].join("");
		expect(testing.rangeFromCells(xml)).toEqual({
			startRow: 2,
			startColumn: 2,
			endRow: 1_048_575,
			endColumn: 16_383,
		});
	});

	test("normalizes relationship targets without escaping the package", () => {
		expect(testing.normalizeWorkbookTarget("worksheets/sheet1.xml")).toBe(
			"xl/worksheets/sheet1.xml",
		);
		expect(testing.normalizeWorkbookTarget("/xl/tables/table1.xml")).toBe(
			"xl/tables/table1.xml",
		);
		expect(testing.normalizeWorkbookTarget("../outside.xml")).toBe(
			"outside.xml",
		);
	});

	test("classifies date styles without shifting time-only formats", () => {
		expect(testing.isDateFormat(14, undefined)).toBe(true);
		expect(testing.isDateFormat(46, undefined)).toBe(false);
		expect(testing.isDateFormat(164, 'yyyy-mm-dd "at" hh:mm')).toBe(true);
		expect(testing.isDateFormat(165, "[h]:mm:ss")).toBe(false);
		expect(testing.isDateFormat(166, "0.00%")).toBe(false);
	});
});

describe("archive safety", () => {
	function archive(options: {
		entries?: number;
		uncompressedSize?: number;
		filenameLength?: number;
	}): Uint8Array {
		const bytes = Buffer.alloc(68);
		bytes.writeUInt32LE(0x02014b50, 0);
		bytes.writeUInt32LE(options.uncompressedSize ?? 0, 24);
		bytes.writeUInt16LE(options.filenameLength ?? 0, 28);
		bytes.writeUInt32LE(0x06054b50, 46);
		bytes.writeUInt16LE(options.entries ?? 1, 56);
		bytes.writeUInt32LE(0, 62);
		return bytes;
	}

	test("rejects missing, malformed, ZIP64, oversized, and truncated directories", () => {
		expect(() => testing.validateZipSizes(new Uint8Array())).toThrow(
			"missing ZIP directory",
		);
		expect(() => testing.validateZipSizes(archive({ entries: 2 }))).toThrow(
			"malformed ZIP directory",
		);
		expect(() =>
			testing.validateZipSizes(archive({ uncompressedSize: 0xffffffff })),
		).toThrow("ZIP64 workbooks are not supported");
		expect(() =>
			testing.validateZipSizes(
				archive({ uncompressedSize: 256 * 1024 * 1024 + 1 }),
			),
		).toThrow("expands beyond the 256 MiB safety limit");
		expect(() =>
			testing.validateZipSizes(archive({ filenameLength: 100 })),
		).toThrow("truncated ZIP directory entry");
	});
});
