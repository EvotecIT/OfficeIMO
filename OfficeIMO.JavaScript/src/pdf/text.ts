import { OfficeIMOError } from "../core/errors.js";
import { ExportCell, assertExportValue } from "../core/presentation.js";
import { checkAbort, pause, taskYieldDue } from "../core/iteration.js";
import { isDate } from "../core/dates.js";
import { defaultBooleanText } from "../core/value-text.js";
import type { ExportValue, ColumnValueContext } from "../core/index.js";
import type { PdfOptions } from "./types.js";
import type { PdfFontResource } from "./font-resources.js";
import type { ExportBudget } from "../core/limits.js";

/** @internal Charge displayed characters: tabs paint four spaces, while CR/LF only separate lines.
 * Raw text remains authoritative for input limits and layout. Skip normalization when no text ceiling observes it. */
export function pdfBudgetText(text: string, budget: ExportBudget): string {
  return budget.limits.maxTextCharacters === undefined ? text : text.replace(/[\t\r\n]/g, character => character === "\t" ? "    " : "");
}

export function synchronousText(value: unknown, name: string): string {
  if (typeof value === "string") return value;
  if (value && typeof (value as { then?: unknown }).then === "function") void Promise.resolve(value).catch(() => {});
  throw new TypeError(name + " must return a synchronous string.");
}
export function displayText(value: ExportValue, options: PdfOptions, context: ColumnValueContext): string {
  assertExportValue(value);
  if (value instanceof ExportCell && value.text !== undefined) return value.text;
  const raw = value instanceof ExportCell ? value.value : value;
  if (options.formatValue) return synchronousText(options.formatValue(raw, context), "formatValue");
  return raw == null ? "" : isDate(raw) ? Number.isFinite(raw.getTime()) ? raw.toISOString() : "" : typeof raw === "number" && !Number.isFinite(raw) ? "" : typeof raw === "boolean" ? (options.booleanText ?? defaultBooleanText)[raw ? "true" : "false"] : String(raw);
}
export interface PdfTextLine { readonly text: string; readonly width: number; }
/** Greedy word boundaries, then scalar boundaries. No character, whitespace or long-word truncation. */
export function wrapText(text: string, font: PdfFontResource, size: number, width: number, maxCharacters: number, maxLines: number, wrap = true): readonly PdfTextLine[] {
  const scanner = wrapSteps(text, font, size, width, maxCharacters, maxLines, wrap);
  while (true) { const next = scanner.next(); if (next.done) return next.value; }
}
/** Let the host process input and cancellation while a large cell or paragraph is laid out. */
export async function wrapTextCooperatively(text: string, font: PdfFontResource, size: number, width: number,
  maxCharacters: number, maxLines: number, wrap = true, signal?: AbortSignal): Promise<readonly PdfTextLine[]> {
  checkAbort(signal);
  const scanner = wrapSteps(text, font, size, width, maxCharacters, maxLines, wrap);
  while (true) {
    const next = scanner.next();
    if (next.done) return next.value;
    if (taskYieldDue()) await pause();
    checkAbort(signal);
  }
}
/** One scanner owns wrapping semantics for synchronous measurements and cooperative layout. */
function* wrapSteps(text: string, font: PdfFontResource, size: number, width: number, maxCharacters: number,
  maxLines: number, wrap: boolean): Generator<void, readonly PdfTextLine[], undefined> {
  if (text.length > maxCharacters) throw new OfficeIMOError("RESOURCE_LIMIT", "maxCellCharacters exceeded.");
  const lines: PdfTextLine[] = [];
  let scanned = 0;
  function emit(chars: readonly string[], measure: number): void {
    if (lines.length >= maxLines) throw new OfficeIMOError("RESOURCE_LIMIT", "maxRowLines exceeded.");
    lines.push({ text: chars.join(""), width: measure });
  }
  for (const paragraph of text.replace(/\r\n?/g, "\n").replace(/\t/g, "    ").split("\n")) {
    let chars: string[] = [], advances: number[] = [], measured = 0, breakAt = -1;
    for (const character of paragraph) {
      if ((++scanned & 4095) === 0) yield;
      const advance = font.width(character.codePointAt(0)!) * size / 1000;
      if (advance > width + .001) throw new RangeError("PDF column is too narrow for a glyph. Increase page/column width or reduce padding/font size.");
      if (measured + advance > width + .001) {
        if (!wrap) throw new RangeError("Unwrapped PDF cell exceeds its column width.");
        const end = breakAt >= 0 ? breakAt + 1 : chars.length;
        const used = advances.slice(0, end).reduce((n, a) => n + a, 0);
        emit(chars.slice(0, end), used);
        chars = chars.slice(end); advances = advances.slice(end); measured = advances.reduce((n, a) => n + a, 0); breakAt = -1;
        for (let i = 0; i < chars.length; i++) if (chars[i] === " ") breakAt = i;
        // A trailing word can still exceed the next line after a prior word boundary.
        if (measured + advance > width + .001) { emit(chars, measured); chars = []; advances = []; measured = 0; breakAt = -1; }
      }
      chars.push(character); advances.push(advance); measured += advance;
      if (character === " ") breakAt = chars.length - 1;
    }
    emit(chars, measured);
    if ((++scanned & 4095) === 0) yield;
  }
  return lines;
}
