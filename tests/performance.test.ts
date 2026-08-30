import { expect, test } from "bun:test";
import { join } from "node:path";
import { WorkbookReader } from "../src/lib.js";

const enabled = process.env.XLSX_TEST_PERFORMANCE === "1" ? test : test.skip;

enabled("formula-heavy financial model stays within a generous budget", () => {
	const started = performance.now();
	const reader = new WorkbookReader(
		join(
			import.meta.dir,
			"fixtures/workbooks/29-advanced-financial-model.xlsx",
		),
		"computed",
	);
	for (const sheet of reader.sheets) reader.toCsv(sheet, 100_000, "none");
	const elapsed = performance.now() - started;
	expect(elapsed).toBeLessThan(5_000);
	expect(process.memoryUsage().rss).toBeLessThan(2 * 1024 * 1024 * 1024);
});

enabled(
	"eight parallel model loads complete within a generous budget",
	async () => {
		const path = join(
			import.meta.dir,
			"fixtures/generated/22-formulas-without-cache.xlsx",
		);
		const started = performance.now();
		await Promise.all(
			Array.from({ length: 8 }, async () => {
				const reader = new WorkbookReader(path, "computed");
				reader.toCsv(
					reader.sheets[0] as NonNullable<(typeof reader.sheets)[0]>,
					100_000,
					"none",
				);
			}),
		);
		expect(performance.now() - started).toBeLessThan(5_000);
	},
);
