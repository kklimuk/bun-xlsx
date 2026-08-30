import { afterEach, describe, expect, test } from "bun:test";
import { mkdtempSync, readdirSync, readFileSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { basename, join } from "node:path";
import { unzipSync, zipSync } from "fflate";
import { PNG } from "pngjs";
import { WorkbookReader } from "../src/lib.js";
import {
	publishRenderedPages,
	renderPdfPages,
	renderWorkbook,
	testing,
} from "../src/render.js";

const temporaryDirectories: string[] = [];

function temporaryDirectory(): string {
	const path = mkdtempSync(join(tmpdir(), "xlsx-render-test-"));
	temporaryDirectories.push(path);
	return path;
}

afterEach(() => {
	delete process.env.XLSX_RENDERER;
	delete process.env.XLSX_SOFFICE;
	for (const path of temporaryDirectories.splice(0))
		rmSync(path, { recursive: true, force: true });
});

describe("PDFium rendering", () => {
	const pdf = join(import.meta.dir, "fixtures/render/two-pages.pdf");

	test("renders page dimensions at the requested DPI", async () => {
		const outDir = temporaryDirectory();
		const paths = await renderPdfPages(pdf, { outDir, dpi: 72 });
		expect(paths.map((path) => basename(path))).toEqual([
			"page-001.png",
			"page-002.png",
		]);
		const first = PNG.sync.read(
			Buffer.from(await Bun.file(paths[0] as string).bytes()),
		);
		const second = PNG.sync.read(
			Buffer.from(await Bun.file(paths[1] as string).bytes()),
		);
		expect([first.width, first.height]).toEqual([200, 100]);
		expect([second.width, second.height]).toEqual([100, 200]);
	});

	test("renders an inclusive page range and rejects out-of-range pages", async () => {
		const outDir = temporaryDirectory();
		const paths = await renderPdfPages(pdf, {
			outDir,
			dpi: 144,
			range: { first: 2, last: 2 },
		});
		expect(paths).toHaveLength(1);
		const image = PNG.sync.read(
			Buffer.from(await Bun.file(paths[0] as string).bytes()),
		);
		expect([image.width, image.height]).toEqual([200, 400]);
		await expect(
			renderPdfPages(pdf, {
				outDir,
				dpi: 72,
				range: { first: 3, last: 3 },
			}),
		).rejects.toThrow("exceeds rendered page count 2");
	});

	test("rejects an invalid PDF without publishing output", async () => {
		const directory = temporaryDirectory();
		const pdfPath = join(directory, "invalid.pdf");
		const outDir = join(directory, "pages");
		await Bun.write(pdfPath, "not a PDF");
		await expect(
			renderPdfPages(pdfPath, { outDir, dpi: 72 }),
		).rejects.toThrow();
		expect(readdirSync(outDir)).toEqual([]);
	});

	test("rejects an oversized page before rasterizing it", async () => {
		const directory = temporaryDirectory();
		const oversized = join(directory, "oversized.pdf");
		const source = await Bun.file(pdf).text();
		await Bun.write(
			oversized,
			source.replace("/MediaBox [0 0 200 100]", "/MediaBox [0 0 999 999]"),
		);
		const outDir = join(directory, "pages");
		await expect(
			renderPdfPages(oversized, {
				outDir,
				dpi: 600,
				range: { first: 1, last: 1 },
			}),
		).rejects.toThrow("pixel page safety limit");
		expect(readdirSync(outDir)).toEqual([]);
	});

	test("rejects an invalid renderer before launching a backend", async () => {
		process.env.XLSX_RENDERER = "invalid";
		const directory = temporaryDirectory();
		await expect(
			renderWorkbook(
				join(import.meta.dir, "fixtures/generated/21-text-csv-escaping.xlsx"),
				{ outDir: directory, dpi: 72 },
			),
		).rejects.toThrow("XLSX_RENDERER must be auto, excel, or libreoffice");
		expect(readdirSync(directory)).toEqual([]);
	});

	test("reports a missing forced renderer without publishing output", async () => {
		process.env.XLSX_RENDERER = "libreoffice";
		process.env.XLSX_SOFFICE = join(temporaryDirectory(), "missing-soffice");
		const output = temporaryDirectory();
		await expect(
			renderWorkbook(
				join(import.meta.dir, "fixtures/generated/21-text-csv-escaping.xlsx"),
				{ outDir: output, dpi: 72 },
			),
		).rejects.toThrow();
		expect(readdirSync(output)).toEqual([]);
	});
});

describe("worksheet selection", () => {
	test("preserves sheet order and makes every unselected sheet very hidden", () => {
		const directory = temporaryDirectory();
		const staged = testing.stageSelectedSheet(
			join(
				import.meta.dir,
				"fixtures/generated/25-hidden-and-very-hidden.xlsx",
			),
			directory,
			"Hidden",
		);
		const workbookXml = new TextDecoder().decode(
			unzipSync(new Uint8Array(readFileSync(staged)))["xl/workbook.xml"],
		);
		const sheets = [...workbookXml.matchAll(/<sheet\b[^>]*\/?\s*>/g)].map(
			(match) => match[0],
		);
		expect(sheets.map((tag) => /name="([^"]*)"/.exec(tag)?.[1])).toEqual([
			"Visible",
			"Hidden",
			"Very Hidden",
		]);
		expect(sheets[1]).not.toContain("state=");
		expect(sheets[0]).toContain('state="veryHidden"');
		expect(sheets[2]).toContain('state="veryHidden"');
		expect(workbookXml).toContain('activeTab="1"');
	});

	test("accepts single-quoted sheet attributes", async () => {
		const directory = temporaryDirectory();
		const source = join(
			import.meta.dir,
			"fixtures/generated/25-hidden-and-very-hidden.xlsx",
		);
		const files = unzipSync(new Uint8Array(readFileSync(source)));
		const workbookName = "xl/workbook.xml";
		const xml = new TextDecoder().decode(files[workbookName]);
		files[workbookName] = new TextEncoder().encode(
			xml.replace(/<sheet\b[^>]*name="Hidden"[^>]*\/>/, (tag) =>
				tag.replaceAll(/="([^"]*)"/g, "='$1'"),
			),
		);
		const quoted = join(directory, "single-quotes.xlsx");
		await Bun.write(quoted, zipSync(files));
		const staged = testing.stageSelectedSheet(quoted, directory, "Hidden");
		const stagedXml = new TextDecoder().decode(
			unzipSync(new Uint8Array(readFileSync(staged)))[workbookName],
		);
		expect(stagedXml).toContain("name='Hidden'");
		expect(stagedXml).not.toContain("name='Hidden' state=");
	});

	test("accepts a worksheet name containing a greater-than sign", async () => {
		const directory = temporaryDirectory();
		const source = join(
			import.meta.dir,
			"fixtures/generated/25-hidden-and-very-hidden.xlsx",
		);
		const files = unzipSync(new Uint8Array(readFileSync(source)));
		const workbookName = "xl/workbook.xml";
		const xml = new TextDecoder().decode(files[workbookName]);
		files[workbookName] = new TextEncoder().encode(
			xml.replace('name="Hidden"', 'name="A>B"'),
		);
		const unusual = join(directory, "greater-than.xlsx");
		await Bun.write(unusual, zipSync(files));
		expect(new WorkbookReader(unusual, "formulas").sheets[1]?.name).toBe("A>B");
		const staged = testing.stageSelectedSheet(unusual, directory, "A>B");
		const stagedXml = new TextDecoder().decode(
			unzipSync(new Uint8Array(readFileSync(staged)))[workbookName],
		);
		expect(stagedXml).toContain('name="A>B"');
		expect(stagedXml).not.toContain('name="A>B" state=');
	});

	test("hides paired as well as self-closing worksheet elements", async () => {
		const directory = temporaryDirectory();
		const source = join(
			import.meta.dir,
			"fixtures/generated/25-hidden-and-very-hidden.xlsx",
		);
		const files = unzipSync(new Uint8Array(readFileSync(source)));
		const workbookName = "xl/workbook.xml";
		const xml = new TextDecoder().decode(files[workbookName]);
		files[workbookName] = new TextEncoder().encode(
			xml.replace(/<sheet\b([^>]*)\/>/g, "<sheet$1></sheet>"),
		);
		const paired = join(directory, "paired-sheets.xlsx");
		await Bun.write(paired, zipSync(files));
		const staged = testing.stageSelectedSheet(paired, directory, "Hidden");
		const stagedXml = new TextDecoder().decode(
			unzipSync(new Uint8Array(readFileSync(staged)))[workbookName],
		);
		expect(stagedXml).toMatch(
			/<sheet\b[^>]*name="Visible"[^>]*state="veryHidden"[^>]*>/,
		);
		expect(stagedXml).toMatch(
			/<sheet\b[^>]*name="Hidden"(?![^>]*state=)[^>]*>/,
		);
	});

	test("rejects oversized expanded archives before inflating them", async () => {
		const directory = temporaryDirectory();
		const oversized = join(directory, "oversized.xlsx");
		const bytes = Buffer.alloc(68);
		bytes.writeUInt32LE(0x02014b50, 0);
		bytes.writeUInt32LE(256 * 1024 * 1024 + 1, 24);
		bytes.writeUInt32LE(0x06054b50, 46);
		bytes.writeUInt16LE(1, 56);
		await Bun.write(oversized, bytes);
		expect(() =>
			testing.stageSelectedSheet(oversized, directory, "Sheet1"),
		).toThrow("expands beyond the 256 MiB safety limit");
	});

	test("rejects an unknown worksheet before launching a renderer", () => {
		expect(() =>
			testing.stageSelectedSheet(
				join(
					import.meta.dir,
					"fixtures/generated/25-hidden-and-very-hidden.xlsx",
				),
				temporaryDirectory(),
				"Missing",
			),
		).toThrow("Unknown worksheet");
	});
});

describe("render publication", () => {
	test("removes stale pages only after replacements are staged", async () => {
		const source = temporaryDirectory();
		const output = temporaryDirectory();
		await Bun.write(join(source, "page-001.png"), "new page");
		await Bun.write(join(output, "page-001.png"), "old page");
		await Bun.write(join(output, "page-002.png"), "stale page");
		await Bun.write(join(output, "notes.txt"), "preserve me");

		const published = await publishRenderedPages(
			[join(source, "page-001.png")],
			output,
		);

		expect(published).toEqual([join(output, "page-001.png")]);
		expect(readdirSync(output).sort()).toEqual(["notes.txt", "page-001.png"]);
		expect(await Bun.file(join(output, "page-001.png")).text()).toBe(
			"new page",
		);
	});

	test("preserves existing pages when staging a replacement fails", async () => {
		const output = temporaryDirectory();
		await Bun.write(join(output, "page-001.png"), "old page");

		await expect(
			publishRenderedPages([join(output, "missing.png")], output),
		).rejects.toThrow();

		expect(await Bun.file(join(output, "page-001.png")).text()).toBe(
			"old page",
		);
	});
});
