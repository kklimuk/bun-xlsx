import { afterEach, expect, test } from "bun:test";
import { existsSync, mkdtempSync, readdirSync, rmSync } from "node:fs";
import { homedir, tmpdir } from "node:os";
import { join } from "node:path";
import { PNG } from "pngjs";
import { renderWorkbook } from "../src/render.js";

const enabled = process.env.XLSX_TEST_EXCEL === "1" ? test : test.skip;
const temporaryDirectories: string[] = [];

afterEach(() => {
	delete process.env.XLSX_RENDERER;
	for (const path of temporaryDirectories.splice(0))
		rmSync(path, { recursive: true, force: true });
});

async function sha256(path: string): Promise<string> {
	const hasher = new Bun.CryptoHasher("sha256");
	hasher.update(await Bun.file(path).bytes());
	return hasher.digest("hex");
}

enabled(
	"Excel exports an upright landscape page without modifying source",
	async () => {
		if (
			process.platform !== "darwin" ||
			!existsSync("/Applications/Microsoft Excel.app")
		)
			throw new Error("XLSX_TEST_EXCEL requires Microsoft Excel on macOS");
		process.env.XLSX_RENDERER = "excel";
		const workbook = join(
			import.meta.dir,
			"fixtures/generated/21-text-csv-escaping.xlsx",
		);
		const output = mkdtempSync(join(tmpdir(), "xlsx-excel-test-"));
		temporaryDirectories.push(output);
		await Bun.write(join(output, "page-999.png"), "stale");
		const beforeHash = await sha256(workbook);
		const container = join(
			homedir(),
			"Library/Containers/com.microsoft.Excel/Data/Documents",
		);
		const stagedBefore = new Set(
			readdirSync(container).filter((name) =>
				name.startsWith(".xlsx-cli-render-"),
			),
		);

		const pages = await renderWorkbook(workbook, { outDir: output, dpi: 72 });

		expect(await sha256(workbook)).toBe(beforeHash);
		expect(readdirSync(output)).toEqual(["page-001.png"]);
		expect(pages).toEqual([join(output, "page-001.png")]);
		const image = PNG.sync.read(
			Buffer.from(await Bun.file(pages[0] as string).bytes()),
		);
		expect(image.width).toBeGreaterThan(image.height);
		const stagedAfter = new Set(
			readdirSync(container).filter((name) =>
				name.startsWith(".xlsx-cli-render-"),
			),
		);
		expect(stagedAfter).toEqual(stagedBefore);
	},
);
