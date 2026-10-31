import { ExportCell, NotSupportedError, OfficeIMOError } from "../../core/index.js";
import type { ExportValue, ExportResult } from "../../core/index.js";
import { inputRows, checkAbort, pause, taskYieldDue } from "../../core/iteration.js";
import { fontProgram } from "../../pdf/font.js";
import { resolveScalar, selectFont, validateFonts } from "../../pdf/font-support.js";
import { settings } from "../../pdf/settings.js";
import type { PdfOptions } from "../../pdf/index.js";
import type { CanopyExport, CanopyExportColumn } from "./types.js";
import { canopyPdfNumbers } from "./pdf-number.js";

/** A bounded lookahead is replayed once; no second capture, full-dataset copy or smaller font is needed. */
export async function withCanopyPdfLayout(source: CanopyExport, captured: readonly CanopyExportColumn[], options: Omit<PdfOptions, "columns">,
  write: (rows: AsyncIterable<readonly ExportValue[]>, layout: Pick<PdfOptions, "columnWidths" | "wideTable">, capturedOptions: Omit<PdfOptions, "columns">) => Promise<ExportResult>): Promise<ExportResult> {
  const prepared = settings({ ...options, columns: source.columns }), { fontSize, padding, page, margins, limits } = prepared;
  options = prepared.options;
  validateFonts(options.fonts, limits.maxFontBytes);
  const available = page.width - margins.left - margins.right;
  const iterator = inputRows(source.rows(), options.signal)[Symbol.asyncIterator](), sample: (readonly ExportValue[])[] = [];
  let done = false, characters = 0;
  try {
    while (sample.length < Math.min(128, options.limits?.maxRows ?? 128) && characters < 262144) {
      const item = await iterator.next();
      if (item.done) { done = true; break; }
      sample.push(item.value);
      characters += item.value.reduce<number>((n, cell) => n + (cell instanceof ExportCell ? cell.text?.length ?? 0 : String(cell ?? "").length), 0);
    }
    const format = canopyPdfNumbers(sample, source.columns), rows = sample.map(format);
    const measure = (text: string, bold = false, italic = false): number => {
      if (text.length > limits.maxCellCharacters) throw new OfficeIMOError("RESOURCE_LIMIT", "maxCellCharacters exceeded.");
      const face = selectFont(options.fonts, bold, italic), program = face.font ? fontProgram(face.font) : undefined;
      let width = 0;
      for (const character of text.replace(/\t/g, "    ")) {
        const scalar = resolveScalar(character.codePointAt(0)!, program, bold);
        if (!scalar.supported) throw new NotSupportedError(scalar.message);
        width += scalar.width * fontSize / 1000;
      }
      return width;
    };
    const largest = (parts: readonly string[], bold = false, italic = false) => parts.reduce((width, part) => Math.max(width, measure(part, bold, italic)), 0);
    let layout: Pick<PdfOptions, "columnWidths" | "wideTable"> = {};
    // Explicit points/bands and shaping providers retain their own layout policy.
    if (!options.columnWidths && typeof options.wideTable !== "object" && !options.textShaper) {
      const minimum = source.columns.map(column => Math.min(available, largest(column.header.split(/\s+/), prepared.options.headerPresentation?.bold) + padding * 2 + .01));
      const natural = source.columns.map((column, index) => Math.min(available, Math.max(minimum[index]!,
        largest(column.header.split(/\r\n?|\n/), prepared.options.headerPresentation?.bold) + padding * 2 + .01)));
      for (const row of rows) {
        for (let i = 0; i < row.length; i++) {
          const cell = row[i], text = cell instanceof ExportCell ? cell.text ?? String(cell.value ?? "") : String(cell ?? "");
          const style = cell instanceof ExportCell ? cell.presentation : undefined;
          const line = largest(text.split(/\r\n?|\n/), style?.bold, style?.italic) + padding * 2 + .01;
          const word = largest(text.split(/\s+/), style?.bold, style?.italic) + padding * 2 + .01;
          natural[i] = Math.min(available, Math.max(natural[i]!, line));
          minimum[i] = Math.min(available, Math.max(minimum[i]!, captured[i]!.kind === "text" ? word : line));
        }
        if (taskYieldDue()) await pause(); checkAbort(options.signal);
      }
      const sum = minimum.reduce((total, width) => total + width, 0), widths = [...minimum];
      if (sum > available + .001) layout = { columnWidths: widths, wideTable: options.wideTable === "reject" ? "reject" : {
        mode: "bands", repeatColumns: widths[0]! + Math.max(...widths.slice(1)) <= available ? [0] : [] } };
      else {
        let spare = available - sum;
        const needs = natural.map((width, index) => Math.max(0, width - widths[index]!)), total = needs.reduce((n, width) => n + width, 0);
        for (let i = 0; i < widths.length; i++) widths[i]! += total ? needs[i]! * Math.min(1, spare / total) : 0;
        spare = available - widths.reduce((n, width) => n + width, 0);
        const text = captured.flatMap((column, index) => column.kind === "text" ? [index] : []);
        for (const i of text.length ? text : widths.map((_, index) => index)) widths[i]! += spare / (text.length || widths.length);
        layout = { columnWidths: widths };
      }
    }
    const replay = async function* () {
      for (const row of rows) yield row;
      while (!done) { const item = await iterator.next(); if (item.done) { done = true; break; } yield format(item.value); }
    };
    return await write(replay(), layout, prepared.options);
  } finally {
    if (!done) { try { void iterator.return(undefined).catch(() => {}); } catch { /* Keep the writer's original failure. */ } }
  }
}
