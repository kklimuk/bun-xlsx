import { describe, expect, test } from "bun:test";
import { mkdtempSync, rmSync } from "node:fs";
import { tmpdir } from "node:os";
import { join, resolve } from "node:path";
import pkg from "../package.json" with { type: "json" };

async function run(
	...args: string[]
): Promise<{ exitCode: number; stdout: string; stderr: string }> {
	const process = Bun.spawn(["bun", "run", "src/cli.ts", ...args], {
		cwd: `${import.meta.dir}/..`,
		stdout: "pipe",
		stderr: "pipe",
	});
	const [exitCode, stdout, stderr] = await Promise.all([
		process.exited,
		new Response(process.stdout).text(),
		new Response(process.stderr).text(),
	]);
	return { exitCode, stdout, stderr };
}

describe("help", () => {
	test("top-level help discovers both commands", async () => {
		const result = await run("--help");
		expect(result.exitCode).toBe(0);
		expect(result.stderr).toBe("");
		expect(result.stdout).toContain("xlsx <command> [options]");
		expect(result.stdout).toContain("read  Write every worksheet");
		expect(result.stdout).toContain("list  Write worksheet names");
		expect(result.stdout).toContain("render  Render workbook print pages");
		expect(result.stdout).toContain(
			"upgrade  Upgrade a standalone release binary",
		);
	});

	test("reports its package version", async () => {
		const result = await run("--version");
		expect(result).toEqual({
			exitCode: 0,
			stdout: `xlsx ${pkg.version}\n`,
			stderr: "",
		});
	});

	test("upgrade has command-specific help without a workbook", async () => {
		const result = await run("upgrade", "--help");
		expect(result.exitCode).toBe(0);
		expect(result.stderr).toBe("");
		expect(result.stdout).toContain(
			"xlsx upgrade — replace a standalone xlsx binary",
		);
		expect(result.stdout).toContain("--dry-run");
	});

	test("package-manager install refuses self-upgrade", async () => {
		const result = await run("upgrade");
		expect(result.exitCode).toBe(1);
		expect(result.stderr).toContain("package-manager-owned");
		expect(result.stderr).toContain("bun add -g @sageling/xlsx");
	});

	test("render has command-specific help without opening the file", async () => {
		const result = await run("render", "--help");
		expect(result.exitCode).toBe(0);
		expect(result.stderr).toBe("");
		expect(result.stdout).toContain(
			"xlsx render — render workbook print pages as images",
		);
		expect(result.stdout).toContain("--pages N[-M]");
		expect(result.stdout).toContain("Microsoft Excel is used when available");
	});

	test("render validates options before opening the file", async () => {
		const result = await run("render", "missing.xlsx", "--dpi", "nope");
		expect(result.exitCode).toBe(1);
		expect(result.stderr).toContain(
			"--dpi must be a number between 36 and 600",
		);
	});

	test("read has command-specific help without opening the file", async () => {
		const result = await run("read", "--help");
		expect(result.exitCode).toBe(0);
		expect(result.stderr).toBe("");
		expect(result.stdout).toContain(
			"xlsx read — write worksheet contents as CSV",
		);
		expect(result.stdout).toContain("--values MODE");
		expect(result.stdout).toContain("exactly two blank lines");
	});

	test("bare invocation shows top-level help", async () => {
		const result = await run();
		expect(result.exitCode).toBe(0);
		expect(result.stderr).toBe("");
		expect(result.stdout).toContain("xlsx <command> [options]");
	});

	test("list has command-specific help without opening the file", async () => {
		const result = await run("list", "--help");
		expect(result.exitCode).toBe(0);
		expect(result.stderr).toBe("");
		expect(result.stdout).toContain(
			"xlsx list — describe workbook worksheets as JSON",
		);
		expect(result.stdout).toContain("No worksheet values are emitted");
		expect(result.stdout).not.toContain("--values MODE");
	});
});

describe("read output", () => {
	const fixture = "tests/fixtures/generated/25-hidden-and-very-hidden.xlsx";

	test("labels multiple worksheet blocks and separates them with two blank lines", async () => {
		const result = await run("read", fixture);
		expect(result.exitCode).toBe(0);
		expect(result.stderr).toBe("");
		expect(result.stdout).toBe(
			"Sheet: Visible\nsheet,visibility\nVisible,visible\n\n\nSheet: Hidden\nhidden value\n\n\nSheet: Very Hidden\nvery hidden value\n",
		);
	});

	test("keeps explicitly selected sheet output as plain CSV", async () => {
		const result = await run("read", fixture, "--sheet", "Hidden");
		expect(result.exitCode).toBe(0);
		expect(result.stderr).toBe("");
		expect(result.stdout).toBe("hidden value\n");
	});

	test("selects raw versus computed values", async () => {
		const formulaFixture =
			"tests/fixtures/generated/22-formulas-without-cache.xlsx";
		const computed = await run("read", formulaFixture);
		const raw = await run("read", formulaFixture, "--values", "raw");
		expect(computed.exitCode).toBe(0);
		expect(computed.stdout).toContain("3,6,3,,87");
		expect(raw.exitCode).toBe(0);
		expect(raw.stdout).toContain("3,=A2*2,=SUM($A$2:A2),,=SUM(B2:B5)+C5");
	});

	test("reads a workbook whose path contains spaces, quotes, and Unicode", async () => {
		const directory = mkdtempSync(join(tmpdir(), "xlsx-path-test-"));
		try {
			const path = join(directory, 'book café "東京".xlsx');
			await Bun.write(
				path,
				Bun.file(resolve("tests/fixtures/generated/21-text-csv-escaping.xlsx")),
			);
			const result = await run("read", path);
			expect(result.exitCode).toBe(0);
			expect(result.stderr).toBe("");
			expect(result.stdout).toContain("naïve café — 東京 — 😀");
		} finally {
			rmSync(directory, { recursive: true, force: true });
		}
	});
});

describe("list output", () => {
	test("writes machine-readable worksheet metadata", async () => {
		const result = await run(
			"list",
			"tests/fixtures/generated/25-hidden-and-very-hidden.xlsx",
		);
		expect(result.exitCode).toBe(0);
		expect(result.stderr).toBe("");
		const output = JSON.parse(result.stdout);
		expect(output.file).toEndWith("25-hidden-and-very-hidden.xlsx");
		expect(output.sheets).toEqual([
			{
				index: 0,
				name: "Visible",
				range: "A1:B2",
				rows: 2,
				columns: 2,
				cells: 4,
			},
			{
				index: 1,
				name: "Hidden",
				range: "A1:A1",
				rows: 1,
				columns: 1,
				cells: 1,
			},
			{
				index: 2,
				name: "Very Hidden",
				range: "A1:A1",
				rows: 1,
				columns: 1,
				cells: 1,
			},
		]);
	});
});

describe("validation and failures", () => {
	const cases: Array<{ args: string[]; message: string }> = [
		{ args: ["wat", "missing.xlsx"], message: "Unknown command: wat" },
		{ args: ["read", "missing.xlsx", "--wat"], message: "Unknown option" },
		{
			args: ["read", "missing.xlsx", "--values", "cached"],
			message: "--values must be computed or raw",
		},
		{
			args: ["read", "missing.xlsx", "--max-cells", "0"],
			message: "--max-cells must be a positive integer",
		},
		{
			args: ["read", "missing.xlsx", "--max-cells", "1.5"],
			message: "--max-cells must be a positive integer",
		},
		{
			args: ["render", "missing.xlsx", "--dpi", "35"],
			message: "--dpi must be a number between 36 and 600",
		},
		{
			args: ["render", "missing.xlsx", "--dpi", "601"],
			message: "--dpi must be a number between 36 and 600",
		},
		{
			args: ["render", "missing.xlsx", "--pages", "0"],
			message: "--pages must be a page number",
		},
		{
			args: ["render", "missing.xlsx", "--pages", "3-2"],
			message: "--pages range must end at or after its start",
		},
		{ args: ["read", "missing.xlsx"], message: "Not a file:" },
		{ args: ["read", "."], message: "Not a file:" },
		{ args: ["missing.xlsx"], message: "Unknown command: missing.xlsx" },
	];

	for (const { args, message } of cases) {
		test(`${args.join(" ")} fails cleanly`, async () => {
			const result = await run(...args);
			expect(result.exitCode).toBe(1);
			expect(result.stdout).toBe("");
			expect(result.stderr).toContain(`xlsx: ${message}`);
		});
	}

	test("reports available worksheets for an unknown sheet", async () => {
		const result = await run(
			"read",
			"tests/fixtures/generated/25-hidden-and-very-hidden.xlsx",
			"--sheet",
			"Missing",
		);
		expect(result.exitCode).toBe(1);
		expect(result.stdout).toBe("");
		expect(result.stderr).toContain(
			'Unknown worksheet "Missing"; available: "Visible", "Hidden", "Very Hidden"',
		);
	});

	test("enforces the cell-output bound without partial stdout", async () => {
		const result = await run(
			"read",
			"tests/fixtures/generated/26-sparse-wide-range.xlsx",
		);
		expect(result.exitCode).toBe(1);
		expect(result.stdout).toBe("");
		expect(result.stderr).toContain("exceeding --max-cells 100000");
	});

	test("checks every sheet bound before emitting the first sheet", async () => {
		const result = await run(
			"read",
			"tests/fixtures/workbooks/30-cross-sheet-dependencies.xlsx",
			"--max-cells",
			"200",
		);
		expect(result.exitCode).toBe(1);
		expect(result.stdout).toBe("");
		expect(result.stderr).toContain(
			'Sheet "Formulas1" covers 238 cells (B2:H35)',
		);
	});
});

describe("concurrency", () => {
	test("parallel computed reads are deterministic", async () => {
		const fixture = "tests/fixtures/generated/22-formulas-without-cache.xlsx";
		const results = await Promise.all(
			Array.from({ length: 8 }, () => run("read", fixture)),
		);
		expect(results.every((result) => result.exitCode === 0)).toBe(true);
		expect(new Set(results.map((result) => result.stdout)).size).toBe(1);
	});
});
