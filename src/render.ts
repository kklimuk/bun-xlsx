import { randomUUID } from "node:crypto";
import {
	existsSync,
	mkdirSync,
	mkdtempSync,
	readdirSync,
	readFileSync,
	renameSync,
	rmSync,
	statSync,
	writeFileSync,
} from "node:fs";
import { homedir, tmpdir } from "node:os";
import { basename, extname, join, resolve } from "node:path";
import { PDFiumLibrary } from "@hyzyla/pdfium";
import pdfiumWasmPath from "@hyzyla/pdfium/dist/pdfium.wasm" with {
	type: "file",
};
import { unzipSync, zipSync } from "fflate";
import { PNG } from "pngjs";
import { attribute, startTags, validateZipSizes } from "./lib.js";
import { MAX_INPUT_BYTES } from "./limits.js";

export type PageRange = { first: number; last: number };

export interface RenderOptions {
	outDir: string;
	dpi: number;
	range?: PageRange;
	sheet?: string;
}

const MAX_RENDERED_PAGES = 200;
const MAX_PAGE_PIXELS = 50_000_000;
const MAX_TOTAL_PIXELS = 500_000_000;
const MAX_TOTAL_PNG_BYTES = 512 * 1024 * 1024;

export async function renderWorkbook(
	workbookPath: string,
	options: RenderOptions,
): Promise<string[]> {
	if (statSync(workbookPath).size > MAX_INPUT_BYTES)
		throw new Error("Workbook exceeds the 256 MiB render-input safety limit");
	const workspace = mkdtempSync(join(tmpdir(), "xlsx-render-"));
	try {
		const renderInput =
			options.sheet !== undefined
				? stageSelectedSheet(workbookPath, workspace, options.sheet)
				: workbookPath;
		const pdfPath = join(workspace, "workbook.pdf");
		const renderer = requestedRenderer();
		if (renderer === "excel") await convertWithExcel(renderInput, pdfPath);
		else await convertWithLibreOffice(renderInput, pdfPath, workspace);
		const stagedPages = await renderPdfPages(pdfPath, {
			dpi: options.dpi,
			range: options.range,
			outDir: join(workspace, "pages"),
		});
		return await publishRenderedPages(stagedPages, options.outDir);
	} finally {
		rmSync(workspace, { recursive: true, force: true });
	}
}

function stageSelectedSheet(
	input: string,
	workspace: string,
	sheet: string,
): string {
	const data = new Uint8Array(readFileSync(input));
	validateZipSizes(data);
	const files = unzipSync(data);
	const name = "xl/workbook.xml";
	const bytes = files[name];
	if (!bytes) throw new Error("Workbook is missing xl/workbook.xml");
	const xml = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
	const tags = startTags(xml, "sheet");
	const names = tags.map((tag) => attribute(tag.attributes, "name") ?? "");
	const selectedIndex = names.indexOf(sheet);
	if (selectedIndex < 0)
		throw new Error(
			`Unknown worksheet ${JSON.stringify(sheet)}; available: ${names.map((value) => JSON.stringify(value)).join(", ")}`,
		);
	if (names.lastIndexOf(sheet) !== selectedIndex)
		throw new Error(
			`Workbook contains duplicate worksheet name ${JSON.stringify(sheet)}`,
		);
	let stagedXml = xml;
	for (let index = tags.length - 1; index >= 0; index -= 1) {
		const tag = tags[index];
		if (!tag) continue;
		const withoutState = tag.text.replace(/\s+state=(?:"[^"]*"|'[^']*')/, "");
		const replacement =
			index === selectedIndex
				? withoutState
				: withoutState.replace(/\s*(\/?>)$/, ' state="veryHidden"$1');
		stagedXml =
			stagedXml.slice(0, tag.start) + replacement + stagedXml.slice(tag.end);
	}
	stagedXml = /\bactiveTab=(?:"[0-9]+"|'[0-9]+')/.test(stagedXml)
		? stagedXml.replace(
				/\bactiveTab=(?:"[0-9]+"|'[0-9]+')/,
				`activeTab="${selectedIndex}"`,
			)
		: stagedXml.replace(
				/<workbookView\b/,
				`<workbookView activeTab="${selectedIndex}"`,
			);
	files[name] = new TextEncoder().encode(stagedXml);
	const staged = join(workspace, "selected-sheet.xlsx");
	writeFileSync(staged, zipSync(files, { level: 6 }), { mode: 0o600 });
	return staged;
}

function requestedRenderer(): "excel" | "libreoffice" {
	const requested = process.env.XLSX_RENDERER ?? "auto";
	if (!new Set(["auto", "excel", "libreoffice"]).has(requested))
		throw new Error("XLSX_RENDERER must be auto, excel, or libreoffice");
	if (requested === "excel") {
		if (!excelAvailable())
			throw new Error("Microsoft Excel renderer is not available");
		return "excel";
	}
	if (requested === "libreoffice") return "libreoffice";
	return excelAvailable() ? "excel" : "libreoffice";
}

function excelAvailable(): boolean {
	return (
		process.platform === "darwin" &&
		existsSync("/Applications/Microsoft Excel.app") &&
		existsSync(excelContainerDirectory())
	);
}

async function convertWithExcel(input: string, output: string): Promise<void> {
	const tag = randomUUID();
	const directory = excelContainerDirectory();
	const stagedInput = join(directory, `.xlsx-cli-render-${tag}.xlsx`);
	const stagedOutput = join(directory, `.xlsx-cli-render-${tag}.pdf`);
	try {
		await Bun.write(stagedInput, Bun.file(input));
		const script = [
			'tell application "Microsoft Excel"',
			"with timeout of 120 seconds",
			`set w to open workbook workbook file name "${appleScriptString(stagedInput)}" update links do not update links read only false`,
			"try",
			"repeat with sheetIndex from 1 to count of sheets of w",
			"set page orientation of page setup object of sheet sheetIndex of w to landscape",
			"end repeat",
			`save workbook as w filename "${appleScriptString(stagedOutput)}" file format PDF file format`,
			"close w saving no",
			"on error messageText number messageNumber",
			"try",
			"close w saving no",
			"end try",
			"error messageText number messageNumber",
			"end try",
			"end timeout",
			"end tell",
		].join("\n");
		const process = Bun.spawn(["osascript", "-e", script], {
			stdout: "pipe",
			stderr: "pipe",
		});
		const [exitCode, stderr] = await Promise.all([
			process.exited,
			new Response(process.stderr).text(),
		]);
		if (exitCode !== 0)
			throw new Error(
				`Microsoft Excel failed with code ${exitCode}: ${stderr.trim() || "no diagnostic output"}`,
			);
		if (!(await Bun.file(stagedOutput).exists()))
			throw new Error(
				"Microsoft Excel reported success but did not create a PDF",
			);
		await Bun.write(output, Bun.file(stagedOutput));
	} finally {
		rmSync(stagedInput, { force: true });
		rmSync(stagedOutput, { force: true });
	}
}

async function convertWithLibreOffice(
	input: string,
	output: string,
	workspace: string,
): Promise<void> {
	const soffice = await findSoffice();
	if (!soffice)
		throw new Error(
			"No renderer found; install Microsoft Excel on macOS or LibreOffice",
		);
	const profile = mkdtempSync(join(tmpdir(), "xlsx-soffice-"));
	try {
		const process = Bun.spawn(
			[
				soffice,
				"--headless",
				`-env:UserInstallation=file://${profile}`,
				"--convert-to",
				"pdf",
				"--outdir",
				workspace,
				input,
			],
			{ stdout: "pipe", stderr: "pipe" },
		);
		let timedOut = false;
		const timeout = setTimeout(() => {
			timedOut = true;
			process.kill();
		}, 120_000);
		const [exitCode, stdout, stderr] = await Promise.all([
			process.exited,
			new Response(process.stdout).text(),
			new Response(process.stderr).text(),
		]);
		clearTimeout(timeout);
		if (timedOut) throw new Error("LibreOffice timed out after 120 seconds");
		if (exitCode !== 0)
			throw new Error(
				`LibreOffice exited with code ${exitCode}: ${stderr.trim() || stdout.trim() || "no diagnostic output"}`,
			);
		const generated = join(workspace, `${basename(input, extname(input))}.pdf`);
		if (!(await Bun.file(generated).exists()))
			throw new Error("LibreOffice reported success but did not create a PDF");
		await Bun.write(output, Bun.file(generated));
	} finally {
		rmSync(profile, { recursive: true, force: true });
	}
}

export async function renderPdfPages(
	pdfPath: string,
	options: { outDir: string; dpi: number; range?: PageRange },
): Promise<string[]> {
	mkdirSync(options.outDir, { recursive: true });
	if (statSync(pdfPath).size > MAX_INPUT_BYTES)
		throw new Error("Rendered PDF exceeds the 256 MiB safety limit");
	const wasmBinary = await Bun.file(
		resolve(import.meta.dir, pdfiumWasmPath),
	).arrayBuffer();
	const library = await PDFiumLibrary.init({ wasmBinary });
	try {
		const document = await library.loadDocument(
			await Bun.file(pdfPath).bytes(),
		);
		try {
			const pageCount = document.getPageCount();
			const first = options.range?.first ?? 1;
			const last = options.range?.last ?? pageCount;
			if (first > pageCount || last > pageCount)
				throw new Error(
					`--pages ${first}-${last} exceeds rendered page count ${pageCount}`,
				);
			const selectedPages = last - first + 1;
			if (selectedPages > MAX_RENDERED_PAGES)
				throw new Error(
					`Rendering ${selectedPages} pages exceeds the ${MAX_RENDERED_PAGES}-page safety limit; use --pages to select a smaller range`,
				);
			const pages = [];
			let totalPixels = 0;
			for (let pageNumber = first; pageNumber <= last; pageNumber += 1) {
				const page = document.getPage(pageNumber - 1);
				pages.push({ page, pageNumber });
				const { originalWidth, originalHeight } = page.getOriginalSize();
				const width = Math.ceil(originalWidth * (options.dpi / 72));
				const height = Math.ceil(originalHeight * (options.dpi / 72));
				const pixels = width * height;
				if (!Number.isSafeInteger(pixels) || pixels > MAX_PAGE_PIXELS)
					throw new Error(
						`Page ${pageNumber} would render ${width}x${height}, exceeding the ${MAX_PAGE_PIXELS.toLocaleString()}-pixel page safety limit`,
					);
				totalPixels += pixels;
			}
			if (totalPixels > MAX_TOTAL_PIXELS)
				throw new Error(
					`Selected pages exceed the ${MAX_TOTAL_PIXELS.toLocaleString()}-pixel render safety limit`,
				);
			const paths: string[] = [];
			let totalPngBytes = 0;
			for (const { page, pageNumber } of pages) {
				const rendered = await page.render({
					scale: options.dpi / 72,
					render: async ({ data, width, height }) =>
						encodePng(data, width, height),
				});
				totalPngBytes += rendered.data.byteLength;
				if (totalPngBytes > MAX_TOTAL_PNG_BYTES)
					throw new Error(
						"Rendered PNG output exceeds the 512 MiB safety limit",
					);
				const path = join(
					options.outDir,
					`page-${String(pageNumber).padStart(3, "0")}.png`,
				);
				await Bun.write(path, rendered.data);
				paths.push(path);
			}
			return paths;
		} finally {
			document.destroy();
		}
	} finally {
		library.destroy();
	}
}

export async function publishRenderedPages(
	stagedPages: string[],
	outDir: string,
): Promise<string[]> {
	mkdirSync(outDir, { recursive: true });
	const stagingDirectory = mkdtempSync(join(outDir, ".xlsx-publish-"));
	const pending: Array<{ temporary: string; destination: string }> = [];
	try {
		for (const [index, staged] of stagedPages.entries()) {
			const name = basename(staged);
			const temporary = join(stagingDirectory, `${index}-${name}`);
			await Bun.write(temporary, Bun.file(staged));
			pending.push({ temporary, destination: join(outDir, name) });
		}
		for (const name of readdirSync(outDir)) {
			if (/^page-[0-9]+\.png$/.test(name)) rmSync(join(outDir, name));
		}
		for (const file of pending) renameSync(file.temporary, file.destination);
		return pending.map((file) => file.destination);
	} finally {
		rmSync(stagingDirectory, { recursive: true, force: true });
	}
}

function encodePng(
	data: Uint8Array,
	width: number,
	height: number,
): Uint8Array {
	const image = new PNG({ width, height });
	image.data = data as unknown as Buffer;
	return PNG.sync.write(image);
}

function excelContainerDirectory(): string {
	return join(
		homedir(),
		"Library/Containers/com.microsoft.Excel/Data/Documents",
	);
}

function appleScriptString(value: string): string {
	return value.replaceAll("\\", "\\\\").replaceAll('"', '\\"');
}

async function findSoffice(): Promise<string | null> {
	if (process.env.XLSX_SOFFICE) return process.env.XLSX_SOFFICE;
	const candidates =
		process.platform === "darwin"
			? [
					"/opt/homebrew/bin/soffice",
					"/Applications/LibreOffice.app/Contents/MacOS/soffice",
				]
			: process.platform === "win32"
				? [
						"C:\\Program Files\\LibreOffice\\program\\soffice.com",
						"C:\\Program Files (x86)\\LibreOffice\\program\\soffice.com",
					]
				: ["/usr/bin/soffice", "/snap/bin/libreoffice"];
	for (const candidate of candidates)
		if (await Bun.file(candidate).exists()) return candidate;
	const lookup = Bun.spawn(
		process.platform === "win32"
			? ["where", "soffice"]
			: ["/bin/sh", "-c", "command -v soffice"],
		{ stdout: "pipe", stderr: "pipe" },
	);
	const output = await new Response(lookup.stdout).text();
	return (await lookup.exited) === 0
		? output.split(/\r?\n/)[0]?.trim() || null
		: null;
}

export const testing = { stageSelectedSheet };
