#!/usr/bin/env bun
import { stat } from "node:fs/promises";
import { basename, extname, resolve } from "node:path";
import { parseArgs } from "node:util";
import pkg from "../package.json" with { type: "json" };
import { LIST_HELP, READ_HELP, RENDER_HELP, TOP_HELP } from "./help.js";
import {
	csvField,
	type LabelMode,
	type SheetInfo,
	type ValueMode,
	WorkbookReader,
} from "./lib.js";
import { type PageRange, renderWorkbook } from "./render.js";

type Command = "read" | "list" | "render";
interface Invocation {
	file: string;
	command: Command;
	args: string[];
}

function invocation(args: string[]): Invocation {
	const candidate = args[0];
	if (candidate !== "read" && candidate !== "list" && candidate !== "render")
		throw new Error(`Unknown command: ${candidate}`);
	const file = args[1];
	if (file === "--help" || file === "-h")
		return { file: "", command: candidate, args: args.slice(1) };
	if (!file) throw new Error(`${candidate} requires a workbook path`);
	return { file, command: candidate, args: args.slice(2) };
}

function parseReadOptions(args: string[]): {
	sheet?: string;
	values: ValueMode;
	labels: LabelMode;
	maxCells: number;
} {
	const parsed = parseArgs({
		args,
		allowPositionals: false,
		strict: true,
		options: {
			sheet: { type: "string" },
			values: { type: "string", default: "all" },
			labels: { type: "string", default: "coordinates" },
			"max-cells": { type: "string", default: "100000" },
			help: { type: "boolean", short: "h" },
		},
	});
	if (parsed.values.help) {
		process.stdout.write(READ_HELP);
		throw new HelpShown();
	}
	const values = parsed.values.values;
	if (values !== "computed" && values !== "formulas" && values !== "all")
		throw new Error("--values must be all, computed, or formulas");
	const labels = parsed.values.labels;
	if (labels !== "coordinates" && labels !== "none")
		throw new Error("--labels must be coordinates or none");
	const maxCells = Number(parsed.values["max-cells"]);
	if (!Number.isSafeInteger(maxCells) || maxCells < 1)
		throw new Error("--max-cells must be a positive integer");
	return { sheet: parsed.values.sheet, values, labels, maxCells };
}

function parseListOptions(args: string[]): void {
	const parsed = parseArgs({
		args,
		allowPositionals: false,
		strict: true,
		options: { help: { type: "boolean", short: "h" } },
	});
	if (parsed.values.help) {
		process.stdout.write(LIST_HELP);
		throw new HelpShown();
	}
}

function parseRenderOptions(args: string[], file: string) {
	const parsed = parseArgs({
		args,
		allowPositionals: false,
		strict: true,
		options: {
			out: { type: "string" },
			sheet: { type: "string" },
			dpi: { type: "string", default: "150" },
			pages: { type: "string" },
			help: { type: "boolean", short: "h" },
		},
	});
	if (parsed.values.help) {
		process.stdout.write(RENDER_HELP);
		throw new HelpShown();
	}
	const dpi = Number(parsed.values.dpi);
	if (!Number.isFinite(dpi) || dpi < 36 || dpi > 600)
		throw new Error("--dpi must be a number between 36 and 600");
	const range = parsed.values.pages
		? parsePageRange(parsed.values.pages)
		: undefined;
	const workbookName = basename(file, extname(file));
	const sheetSuffix = parsed.values.sheet
		? `-${safePathComponent(parsed.values.sheet)}`
		: "";
	return {
		outDir: resolve(
			parsed.values.out ?? `./${workbookName}${sheetSuffix}-pages`,
		),
		dpi,
		range,
		sheet: parsed.values.sheet,
	};
}

function safePathComponent(value: string): string {
	const unsafe = '<>:"/\\|?*';
	return [...value]
		.map((character) =>
			unsafe.includes(character) || character.charCodeAt(0) < 32
				? "_"
				: character,
		)
		.join("");
}

function parsePageRange(value: string): PageRange {
	const match = /^([1-9][0-9]*)(?:-([1-9][0-9]*))?$/.exec(value);
	if (!match)
		throw new Error('--pages must be a page number or range like "1-3"');
	const first = Number(match[1]);
	const last = Number(match[2] ?? match[1]);
	if (last < first)
		throw new Error("--pages range must end at or after its start");
	return { first, last };
}

class HelpShown extends Error {}

async function assertFile(path: string): Promise<void> {
	const info = await stat(path).catch(() => null);
	if (!info?.isFile()) throw new Error(`Not a file: ${path}`);
}

export async function main(args: string[]): Promise<number> {
	try {
		if (args[0] === "--version" || args[0] === "-V") {
			process.stdout.write(`xlsx ${pkg.version}\n`);
			return 0;
		}
		if (args[0] === "upgrade") {
			const { upgrade } = await import("./upgrade.js");
			return await upgrade(args.slice(1));
		}
		if (
			args.length === 0 ||
			args[0] === "--help" ||
			args[0] === "-h" ||
			args[0] === "help"
		) {
			process.stdout.write(TOP_HELP);
			return 0;
		}
		const call = invocation(args);
		const file = resolve(call.file);
		if (call.command === "render") {
			const options = parseRenderOptions(call.args, file);
			await assertFile(file);
			const pages = await renderWorkbook(file, options);
			if (pages.length > 0) process.stdout.write(`${pages.join("\n")}\n`);
			return 0;
		}
		if (call.command === "list") {
			parseListOptions(call.args);
			await assertFile(file);
			const workbook = new WorkbookReader(file, "formulas");
			process.stdout.write(
				`${JSON.stringify({ file, sheets: workbook.sheets }, null, 2)}\n`,
			);
			return 0;
		}
		const options = parseReadOptions(call.args);
		await assertFile(file);
		const workbook = new WorkbookReader(file, options.values);
		let sheets: SheetInfo[] = workbook.sheets;
		if (options.sheet !== undefined) {
			const match = sheets.find((sheet) => sheet.name === options.sheet);
			if (!match)
				throw new Error(
					`Unknown worksheet ${JSON.stringify(options.sheet)}; available: ${sheets.map((sheet) => JSON.stringify(sheet.name)).join(", ")}`,
				);
			sheets = [match];
		}
		const includeSheetNames = sheets.length > 1;
		for (const sheet of sheets)
			workbook.assertCellLimit(sheet, options.maxCells);
		for (const [index, sheet] of sheets.entries()) {
			if (index > 0) process.stdout.write("\n\n");
			if (includeSheetNames)
				process.stdout.write(`Sheet: ${csvField(sheet.name)}\n`);
			process.stdout.write(
				workbook.toCsv(sheet, options.maxCells, options.labels),
			);
		}
		return 0;
	} catch (error) {
		if (error instanceof HelpShown) return 0;
		process.stderr.write(
			`xlsx: ${error instanceof Error ? error.message : String(error)}\n`,
		);
		return 1;
	}
}

if (import.meta.main) process.exit(await main(Bun.argv.slice(2)));
