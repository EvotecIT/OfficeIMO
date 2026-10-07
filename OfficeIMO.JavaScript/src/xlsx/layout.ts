import type { CellValue } from "../core/index.js";
import type { ProjectionColumn } from "../internal/columns.js";
import { ExportCell } from "../core/presentation.js";
import { Cell, cellText, columnName } from "./values.js";
import { escapeOoxmlAttribute } from "../xml/index.js";
import type { InvalidCharacterPolicy } from "../xml/index.js";
import type { SheetOptions, TotalOperation, PrintOptions } from "./types.js";
import { MergeRegions } from "./regions.js";
import { OfficeIMOError } from "../core/errors.js";
import { createTotals } from "../internal/totals.js";
import type { NumericAggregate } from "../internal/totals.js";

/** @internal A computed formula with a numeric cache; no public arbitrary-formula input. */
export class ComputedTotal { constructor(readonly value: number | null, readonly formula: string, readonly operation: TotalOperation) {} }
/** @internal Bounded report layout and incremental numeric aggregates. */
export class ReportLayout {
  readonly headings: readonly (readonly string[])[];
  readonly merges: readonly string[];
  readonly headerRows: number;
  readonly firstHeaderRow: number;
  readonly regions: MergeRegions;
  readonly widths: (number | undefined)[];
  readonly sampleRows: number;
  private readonly aggregates: readonly NumericAggregate[];
  constructor(readonly columns: readonly ProjectionColumn[], readonly options: SheetOptions, policy: InvalidCharacterPolicy, maximumMerges = 10000, maximumSampleCells = 100000) {
    const titleRows = options.title === undefined ? 0 : 1;
    if (options.title !== undefined) {
      if (!columns.length) throw new TypeError("Report titles require declared columns.");
      if (typeof options.title.text !== "string") throw new TypeError("Report title text must be a string.");
      cellText(options.title.text, policy);
      if (options.title.height !== undefined && (!Number.isFinite(options.title.height) || options.title.height <= 0 || options.title.height > 409)) throw new RangeError("Title height must be positive and at most 409 points.");
    }
    const depth = Math.max(0, ...columns.map(c => c.groups?.length ?? 0));
    if (depth > 16) throw new RangeError("Grouped headings support at most 16 levels.");
    if (depth && options.includeHeader === false) throw new TypeError("Grouped headings require leaf headers.");
    const headings: string[][] = [], merges: string[] = [];
    if (titleRows && columns.length > 1) merges.push("A1:" + columnName(columns.length) + "1");
    for (let level = 0; level < depth; level++) {
      const values = columns.map(c => c.groups?.[level] ?? "");
      for (const value of values) cellText(value, policy);
      for (let first = 0; first < columns.length;) {
        let last = first;
        const prefix = JSON.stringify(columns[first]!.groups?.slice(0, level + 1));
        while (last + 1 < columns.length && values[first] && JSON.stringify(columns[last + 1]!.groups?.slice(0, level + 1)) === prefix) last++;
        if (last > first) { merges.push(columnName(first + 1) + (level + 1 + titleRows) + ":" + columnName(last + 1) + (level + 1 + titleRows)); for (let i = first + 1; i <= last; i++) values[i] = ""; }
        first = last + 1;
      }
      headings.push(values);
    }
    this.headerRows = titleRows + (columns.length && options.includeHeader !== false ? depth + 1 : 0);
    this.firstHeaderRow = columns.length && options.includeHeader !== false ? titleRows + 1 : 0;
    if (options.mergedCells !== undefined && !Array.isArray(options.mergedCells)) throw new TypeError("Merged ranges must be an array.");
    if (options.mergedCells && options.mergedCells.length + merges.length > maximumMerges) throw new OfficeIMOError("RESOURCE_LIMIT", "maxMergedRanges exceeded.");
    this.regions = new MergeRegions([...merges, ...(options.mergedCells ?? [])], columns.length, maximumMerges);
    if (options.table) for (const ref of options.mergedCells ?? []) if (Number(ref.split(":")[1]!.match(/\d+$/)![0]) >= this.headerRows) throw new TypeError("Merged ranges cannot intersect a native table.");
    this.headings = headings; this.merges = this.regions.references;
    const sizing = options.autoSize;
    this.sampleRows = sizing ? sizing.sampleRows ?? Math.min(100, Math.floor(maximumSampleCells / Math.max(1, columns.length))) : 0;
    if (!Number.isInteger(this.sampleRows) || this.sampleRows < 0 || this.sampleRows > 10000) throw new RangeError("Width sampling must use from 0 through 10,000 rows.");
    const min = sizing?.minWidth ?? 8, max = sizing?.maxWidth ?? 60;
    if (!Number.isFinite(min) || !Number.isFinite(max) || min < 0 || min > max || max > 255) throw new RangeError("Width bounds must satisfy 0 <= minWidth <= maxWidth <= 255.");
    this.widths = columns.map(c => c.width ?? (sizing ? Math.max(min, Math.min(max, c.header.length + 2)) : undefined));
    if (options.footer?.values && options.footer.values.length > columns.length) throw new RangeError("Footer has more values than declared columns.");
    this.aggregates = createTotals(columns, options.footer?.totals);
    if (options.print) validatePrint(options.print, policy);
  }
  sample(values: readonly unknown[]): void {
    const min = this.options.autoSize?.minWidth ?? 8, max = this.options.autoSize?.maxWidth ?? 60;
    values.forEach((value, i) => {
      if (!this.columns[i] || this.columns[i]!.width !== undefined) return;
      const raw = value instanceof ExportCell ? value.text ?? value.value : value instanceof Cell ? value.value : value;
      const text = raw instanceof Date ? Number.isFinite(raw.getTime()) ? raw.toISOString() : "" : raw == null ? "" : String(raw);
      // Inspect at most a small prefix. A single long value cannot make a huge sampling allocation.
      const length = Math.max(0, ...text.slice(0, Math.ceil(max * 4)).split(/\r?\n/).map(line => [...line].length));
      this.widths[i] = Math.max(this.widths[i] ?? min, Math.min(max, length + 2));
    });
  }
  accept(index: number, value: CellValue): void {
    this.aggregates[index]!.accept(value);
  }
  footer(rows: number): readonly unknown[] {
    return this.columns.map((_, i) => {
      const total = this.aggregates[i]!, operation = total.operation;
      if (!operation) return this.options.footer?.values?.[i];
      const value = total.value();
      return new ComputedTotal(value, totalFormula(operation, columnName(i + 1), this.headerRows, rows), operation);
    });
  }
}
/** @internal Worksheet formulas and native table metadata must describe the same calculation. */
export function totalFormula(operation: TotalOperation, letter: string, headerRows: number, rows: number): string {
  const code = { sum: 109, count: 102, average: 101, min: 105, max: 104 }[operation];
  const guarded = ["average", "min", "max"].includes(operation);
  const range = letter + (headerRows + 1) + ":" + letter + (headerRows + rows), subtotal = "SUBTOTAL(" + code + "," + range + ")";
  return !rows ? guarded ? '""' : "0" : guarded ? 'IF(SUBTOTAL(102,' + range + ')=0,"",' + subtotal + ')' : subtotal;
}
function validatePrint(options: PrintOptions, policy: InvalidCharacterPolicy): void {
  if (options.paper !== undefined && !["A4", "Letter"].includes(options.paper)) throw new TypeError("Print paper must be A4 or Letter.");
  if (options.orientation !== undefined && !["portrait", "landscape"].includes(options.orientation)) throw new TypeError("Invalid print orientation.");
  for (const value of [options.fitToWidth, options.fitToHeight]) if (value !== undefined && (!Number.isInteger(value) || value < 0 || value > 32767)) throw new RangeError("Print fit dimensions must be integers from 0 through 32,767.");
  for (const value of Object.values(options.margins ?? {})) if (!Number.isFinite(value) || value < 0 || value > 10) throw new RangeError("Print margins must be from 0 through 10 inches.");
  for (const text of [options.header, options.footer]) if (text !== undefined) {
    if (typeof text !== "string" || text.length > 250) throw new TypeError("Print header/footer must be strings of at most 250 UTF-16 units.");
    cellText(text, policy);
    if (("&C" + text.replace(/&/g, "&&")).length > 255) throw new RangeError("Encoded print header/footer exceeds Excel's 255-character limit.");
  }
}
/** @internal XML follows the worksheet schema's pageMargins/pageSetup/headerFooter order. */
export function printXml(options: PrintOptions | undefined, policy: InvalidCharacterPolicy): string {
  if (!options) return "";
  const margins = { left: .25, right: .25, top: .5, bottom: .5, header: .2, footer: .2, ...options.margins };
  return '<pageMargins' + Object.entries(margins).map(([key, value]) => ' ' + key + '="' + value + '"').join("") + '/>' +
    '<pageSetup paperSize="' + (options.paper === "Letter" ? 1 : 9) + '" orientation="' + (options.orientation ?? "landscape") + '" fitToWidth="' + (options.fitToWidth ?? 1) + '" fitToHeight="' + (options.fitToHeight ?? 0) + '"/>' +
    (options.header !== undefined || options.footer !== undefined ? '<headerFooter>' +
      (options.header === undefined ? "" : '<oddHeader>' + escapeOoxmlAttribute("&C" + options.header.replace(/&/g, "&&"), policy) + '</oddHeader>') +
      (options.footer === undefined ? "" : '<oddFooter>' + escapeOoxmlAttribute("&C" + options.footer.replace(/&/g, "&&"), policy) + '</oddFooter>') + '</headerFooter>' : "");
}
