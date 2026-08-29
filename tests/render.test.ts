import { afterEach, describe, expect, test } from "bun:test";
import { mkdtempSync, readdirSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { basename, join } from "node:path";
import { PNG } from "pngjs";
import {
	publishRenderedPages,
	renderPdfPages,
	renderWorkbook,
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
