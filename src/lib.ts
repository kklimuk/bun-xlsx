import {
	mkdtempSync,
	readFileSync,
	rmSync,
	statSync,
	writeFileSync,
} from "node:fs";
import { tmpdir } from "node:os";
import { join, posix } from "node:path";
import { Model } from "@ironcalc/nodejs";
import { unzipSync, zipSync } from "fflate";
import { MAX_INPUT_BYTES } from "./limits.js";

export type ValueMode = "computed" | "formulas" | "all";
export type LabelMode = "coordinates" | "none";

export interface SheetInfo {
	index: number;
	name: string;
	range: string | null;
	rows: number;
	columns: number;
	cells: number;
}

type SheetProperties = { name: string; state: string; sheet_id: number };
type Range = {
	startRow: number;
	startColumn: number;
	endRow: number;
	endColumn: number;
};

export class WorkbookReader {
	readonly sheets: SheetInfo[];
	private readonly model: Model;
	private readonly ranges: Array<Range | null>;

	constructor(
		path: string,
		private readonly mode: ValueMode,
	) {
		if (statSync(path).size > MAX_INPUT_BYTES)
			throw new Error(
				"Workbook exceeds the 256 MiB compressed-size safety limit",
			);
		const data = readFileSync(path);
		validateZipSizes(data);
		const files = unzipSync(data);
		const structuralRanges = readSheetRanges(files);
		this.model = loadIronCalcModel(path, files);
		if (this.mode !== "formulas") this.model.evaluate();

		const properties =
			this.model.getWorksheetsProperties() as SheetProperties[];
		this.ranges = properties.map(({ name }) => {
			if (!structuralRanges.has(name)) {
				throw new Error(
					`The workbook readers disagreed about worksheet ${JSON.stringify(name)}`,
				);
			}
			return structuralRanges.get(name) ?? null;
		});
		this.sheets = properties.map((property, index) => {
			const range = this.ranges[index];
			const rows = range ? range.endRow - range.startRow + 1 : 0;
			const columns = range ? range.endColumn - range.startColumn + 1 : 0;
			return {
				index,
				name: property.name,
				range: range ? encodeRange(range) : null,
				rows,
				columns,
				cells: rows * columns,
			};
		});
	}

	toCsv(
		sheet: SheetInfo,
		maxCells: number,
		labels: LabelMode = "coordinates",
	): string {
		this.assertCellLimit(sheet, maxCells);
		const range = this.ranges[sheet.index];
		if (!range) return "";

		const lines: string[] = [];
		if (labels === "coordinates") {
			const columns = ["Row"];
			for (
				let column = range.startColumn;
				column <= range.endColumn;
				column += 1
			)
				columns.push(columnName(column));
			lines.push(columns.join(","));
		}
		for (let row = range.startRow; row <= range.endRow; row += 1) {
			const fields: string[] =
				labels === "coordinates" ? [String(row + 1)] : [];
			for (
				let column = range.startColumn;
				column <= range.endColumn;
				column += 1
			) {
				const value = this.cellValue(sheet.index, row + 1, column + 1);
				fields.push(csvField(value));
			}
			lines.push(fields.join(","));
		}
		return `${lines.join("\n")}\n`;
	}

	private cellValue(sheet: number, row: number, column: number): string {
		if (this.mode === "formulas")
			return this.model.getCellContent(sheet, row, column);
		const computed = this.model.getFormattedCellValue(sheet, row, column);
		if (this.mode === "computed") return computed;
		const formula = this.model.getCellFormula(sheet, row, column);
		return formula === null ? computed : `${computed}⟦${formula}⟧`;
	}

	assertCellLimit(sheet: SheetInfo, maxCells: number): void {
		if (sheet.cells > maxCells) {
			throw new Error(
				`Sheet ${JSON.stringify(sheet.name)} covers ${sheet.cells} cells (${sheet.range}), exceeding --max-cells ${maxCells}`,
			);
		}
	}
}

const MAX_ARCHIVE_BYTES = 256 * 1024 * 1024;

function readSheetRanges(
	files: Record<string, Uint8Array>,
): Map<string, Range | null> {
	const workbook = textFile(files, "xl/workbook.xml");
	const relationships = textFile(files, "xl/_rels/workbook.xml.rels");
	const targets = new Map<string, string>();
	for (const { attributes } of startTags(relationships, "Relationship")) {
		const id = attribute(attributes, "Id");
		const target = attribute(attributes, "Target");
		if (id && target && !target.includes(":"))
			targets.set(id, normalizeWorkbookTarget(target));
	}

	const result = new Map<string, Range | null>();
	for (const { attributes } of startTags(workbook, "sheet")) {
		const name = attribute(attributes, "name");
		const relationshipId = attribute(attributes, "r:id");
		if (!name || !relationshipId) continue;
		const target = targets.get(relationshipId);
		if (!target?.startsWith("xl/worksheets/")) continue;
		const xml = textFile(files, target);
		const dimension = /<dimension\b[^>]*\bref="([^"]+)"/.exec(xml)?.[1];
		const contentRange = rangeFromCells(xml);
		const hasUnaddressedContent =
			contentRange === null && /<(?:v|f|is)\b/.test(xml);
		result.set(
			name,
			contentRange ??
				(hasUnaddressedContent && dimension ? decodeRange(dimension) : null),
		);
	}
	return result;
}

function loadIronCalcModel(
	path: string,
	files: Record<string, Uint8Array>,
): Model {
	if (!normalizeForIronCalc(files))
		return Model.fromXlsx(path, "en", "UTC", "en");
	const directory = mkdtempSync(join(tmpdir(), "xlsx-ironcalc-"));
	const normalizedPath = join(directory, "normalized.xlsx");
	try {
		writeFileSync(normalizedPath, zipSync(files, { level: 6 }), {
			mode: 0o600,
		});
		return Model.fromXlsx(normalizedPath, "en", "UTC", "en");
	} finally {
		rmSync(directory, { recursive: true, force: true });
	}
}

function normalizeForIronCalc(files: Record<string, Uint8Array>): boolean {
	let changed = normalizeAbsoluteRelationships(files);
	const workbookName = "xl/workbook.xml";
	const workbook = textFile(files, workbookName);
	if (!/<workbookPr\b[^>]*\bdate1904="(?:1|true)"/i.test(workbook))
		return changed;

	const dateStyles = dateStyleIndexes(textFile(files, "xl/styles.xml"));
	for (const [name, bytes] of Object.entries(files)) {
		if (!/^xl\/worksheets\/[^/]+\.xml$/.test(name)) continue;
		const original = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
		const normalized = shift1904DateCells(original, dateStyles);
		if (normalized !== original) {
			files[name] = new TextEncoder().encode(normalized);
			changed = true;
		}
	}
	files[workbookName] = new TextEncoder().encode(
		workbook.replace(
			/(<workbookPr\b[^>]*\bdate1904=")(?:1|true)(")/i,
			(_match, prefix: string, suffix: string) => `${prefix}0${suffix}`,
		),
	);
	return true;
}

function normalizeAbsoluteRelationships(
	files: Record<string, Uint8Array>,
): boolean {
	let changed = false;
	for (const [name, bytes] of Object.entries(files)) {
		if (!name.endsWith(".rels")) continue;
		const original = new TextDecoder("utf-8", { fatal: true }).decode(bytes);
		const sourceDirectory = relationshipSourceDirectory(name);
		const normalized = original.replace(
			/<Relationship\b[^>]*\/?>/g,
			(relationship) => {
				if (/\bTargetMode="External"/i.test(relationship)) return relationship;
				return relationship.replace(
					/\bTarget="\/([^"]+)"/,
					(_attribute, absolute: string) =>
						`Target="${posix.relative(sourceDirectory || ".", absolute)}"`,
				);
			},
		);
		if (normalized !== original) {
			files[name] = new TextEncoder().encode(normalized);
			changed = true;
		}
	}
	return changed;
}

function relationshipSourceDirectory(name: string): string {
	const marker = "/_rels/";
	const markerIndex = name.lastIndexOf(marker);
	if (markerIndex >= 0) return name.slice(0, markerIndex);
	if (name === "_rels/.rels") return "";
	return posix.dirname(name);
}

function dateStyleIndexes(styles: string): Set<number> {
	const customFormats = new Map<number, string>();
	for (const match of styles.matchAll(/<numFmt\b([^>]*)\/?>/g)) {
		const attributes = match[1];
		if (attributes === undefined) continue;
		const id = Number(attribute(attributes, "numFmtId"));
		const code = attribute(attributes, "formatCode");
		if (Number.isInteger(id) && code !== null) customFormats.set(id, code);
	}

	const indexes = new Set<number>();
	const cellXfs = /<cellXfs\b[^>]*>([\s\S]*?)<\/cellXfs>/.exec(styles)?.[1];
	if (!cellXfs) return indexes;
	let index = 0;
	for (const match of cellXfs.matchAll(/<xf\b([^>]*)\/?>/g)) {
		const attributes = match[1];
		if (attributes !== undefined) {
			const id = Number(attribute(attributes, "numFmtId") ?? 0);
			if (isDateFormat(id, customFormats.get(id))) indexes.add(index);
		}
		index += 1;
	}
	return indexes;
}

function isDateFormat(id: number, custom: string | undefined): boolean {
	if (
		(id >= 14 && id <= 17) ||
		id === 22 ||
		(id >= 27 && id <= 36) ||
		(id >= 50 && id <= 58)
	)
		return true;
	if (id >= 45 && id <= 47) return false;
	if (!custom) return false;
	const visible = custom
		.replace(/"[^"]*"/g, "")
		.replace(/\\./g, "")
		.replace(/\[[^\]]*]/g, "")
		.toLowerCase();
	return /[yd]/.test(visible);
}

function shift1904DateCells(
	worksheet: string,
	dateStyles: Set<number>,
): string {
	return worksheet.replace(
		/<c\b([^>]*)>([\s\S]*?)<\/c>/g,
		(cell, attributes: string, body: string) => {
			const style = Number(attribute(attributes, "s") ?? 0);
			if (!dateStyles.has(style) || /<f\b/.test(body)) return cell;
			return cell.replace(
				/(<v>)([^<]+)(<\/v>)/,
				(original, open: string, number: string, close: string) => {
					const serial = Number(number);
					return Number.isFinite(serial)
						? `${open}${serial + 1462}${close}`
						: original;
				},
			);
		},
	);
}

export function validateZipSizes(data: Uint8Array): void {
	const minimum = Math.max(0, data.length - 65_557);
	let end = -1;
	for (let offset = data.length - 22; offset >= minimum; offset -= 1) {
		if (readU32(data, offset) === 0x06054b50) {
			end = offset;
			break;
		}
	}
	if (end < 0) throw new Error("Invalid XLSX: missing ZIP directory");
	const entries = readU16(data, end + 10);
	let offset = readU32(data, end + 16);
	let total = 0;
	for (let entry = 0; entry < entries; entry += 1) {
		if (offset + 46 > data.length || readU32(data, offset) !== 0x02014b50) {
			throw new Error("Invalid XLSX: malformed ZIP directory");
		}
		const size = readU32(data, offset + 24);
		if (size === 0xffffffff)
			throw new Error("ZIP64 workbooks are not supported");
		total += size;
		if (size > MAX_ARCHIVE_BYTES || total > MAX_ARCHIVE_BYTES) {
			throw new Error(
				`Workbook expands beyond the ${MAX_ARCHIVE_BYTES / 1024 / 1024} MiB safety limit`,
			);
		}
		const filenameLength = readU16(data, offset + 28);
		const extraLength = readU16(data, offset + 30);
		const commentLength = readU16(data, offset + 32);
		offset += 46 + filenameLength + extraLength + commentLength;
		if (offset > end)
			throw new Error("Invalid XLSX: truncated ZIP directory entry");
	}
}

function readU16(data: Uint8Array, offset: number): number {
	const low = data[offset];
	const high = data[offset + 1];
	if (low === undefined || high === undefined) {
		throw new Error("Invalid XLSX: truncated ZIP directory");
	}
	return low | (high << 8);
}

function readU32(data: Uint8Array, offset: number): number {
	return (readU16(data, offset) | (readU16(data, offset + 2) << 16)) >>> 0;
}

function textFile(files: Record<string, Uint8Array>, name: string): string {
	const value = files[name];
	if (!value) throw new Error(`Invalid XLSX: missing ${name}`);
	return new TextDecoder("utf-8", { fatal: true }).decode(value);
}

export function attribute(source: string, name: string): string | null {
	const escaped = name.replace(/[.*+?^${}()|[\]\\]/g, "\\$&");
	const match = new RegExp(`(?:^|\\s)${escaped}=(?:"([^"]*)"|'([^']*)')`).exec(
		source,
	);
	const value = match?.[1] ?? match?.[2];
	return value === undefined ? null : decodeXml(value);
}

export interface XmlStartTag {
	attributes: string;
	end: number;
	start: number;
	text: string;
}

export function startTags(source: string, name: string): XmlStartTag[] {
	const tags: XmlStartTag[] = [];
	const prefix = `<${name}`;
	let searchFrom = 0;
	while (searchFrom < source.length) {
		const start = source.indexOf(prefix, searchFrom);
		if (start < 0) break;
		const boundary = source[start + prefix.length];
		if (boundary && !/[\s/>]/.test(boundary)) {
			searchFrom = start + prefix.length;
			continue;
		}
		let quote = "";
		let end = start + prefix.length;
		for (; end < source.length; end += 1) {
			const character = source[end];
			if ((character === '"' || character === "'") && !quote) quote = character;
			else if (character === quote) quote = "";
			else if (character === ">" && !quote) break;
		}
		if (end >= source.length)
			throw new Error(`Invalid XLSX: unterminated <${name}> tag`);
		const text = source.slice(start, end + 1);
		tags.push({
			attributes: text.slice(prefix.length, -1),
			end: end + 1,
			start,
			text,
		});
		searchFrom = end + 1;
	}
	return tags;
}

export function decodeXml(value: string): string {
	return value.replace(
		/&(?:#x([0-9a-f]+)|#([0-9]+)|amp|lt|gt|quot|apos);/gi,
		(entity, hex: string, decimal: string) => {
			if (hex) return String.fromCodePoint(Number.parseInt(hex, 16));
			if (decimal) return String.fromCodePoint(Number.parseInt(decimal, 10));
			return (
				(
					{
						"&amp;": "&",
						"&lt;": "<",
						"&gt;": ">",
						"&quot;": '"',
						"&apos;": "'",
					} as Record<string, string>
				)[entity.toLowerCase()] ?? entity
			);
		},
	);
}

function normalizeWorkbookTarget(target: string): string {
	const parts: string[] = [];
	const rooted = target.startsWith("/") ? target.slice(1) : `xl/${target}`;
	for (const part of rooted.replaceAll("\\", "/").split("/")) {
		if (!part || part === ".") continue;
		if (part === "..") parts.pop();
		else parts.push(part);
	}
	return parts.join("/");
}

function rangeFromCells(xml: string): Range | null {
	let range: Range | null = null;
	for (const match of xml.matchAll(
		/<c\b[^>]*\br="([A-Z]+[0-9]+)"[^>]*>([\s\S]*?)<\/c>/g,
	)) {
		const reference = match[1];
		const body = match[2];
		if (reference === undefined || body === undefined) continue;
		if (!/<(?:v|f|is)\b/.test(body)) continue;
		const cell = decodeCell(reference);
		range = range
			? {
					startRow: Math.min(range.startRow, cell.row),
					startColumn: Math.min(range.startColumn, cell.column),
					endRow: Math.max(range.endRow, cell.row),
					endColumn: Math.max(range.endColumn, cell.column),
				}
			: {
					startRow: cell.row,
					startColumn: cell.column,
					endRow: cell.row,
					endColumn: cell.column,
				};
	}
	return range;
}

function decodeRange(reference: string): Range {
	const [start, end = start] = reference.replaceAll("$", "").split(":");
	if (!start || !end) throw new Error(`Invalid worksheet range: ${reference}`);
	const a = decodeCell(start);
	const b = decodeCell(end);
	return {
		startRow: a.row,
		startColumn: a.column,
		endRow: b.row,
		endColumn: b.column,
	};
}

function decodeCell(reference: string): { row: number; column: number } {
	const match = /^([A-Z]+)([1-9][0-9]*)$/.exec(reference);
	if (!match) throw new Error(`Invalid worksheet range: ${reference}`);
	const letters = match[1];
	const row = match[2];
	if (!letters || !row)
		throw new Error(`Invalid worksheet range: ${reference}`);
	let column = 0;
	for (const character of letters)
		column = column * 26 + character.charCodeAt(0) - 64;
	return { row: Number(row) - 1, column: column - 1 };
}

function encodeRange(range: Range): string {
	const cell = (row: number, column: number) =>
		`${columnName(column)}${row + 1}`;
	return `${cell(range.startRow, range.startColumn)}:${cell(range.endRow, range.endColumn)}`;
}

function columnName(column: number): string {
	let letters = "";
	for (let value = column + 1; value > 0; value = Math.floor((value - 1) / 26))
		letters = String.fromCharCode(((value - 1) % 26) + 65) + letters;
	return letters;
}

export function csvField(value: string): string {
	if (!/[",\r\n]/.test(value)) return value;
	return `"${value.replaceAll('"', '""')}"`;
}

export const testing = {
	columnName,
	decodeRange,
	decodeXml,
	isDateFormat,
	normalizeWorkbookTarget,
	rangeFromCells,
	startTags,
	validateZipSizes,
};
