import { OfficeIMOError } from "./errors.js";
import type { ByteSink } from "./sinks.js";

/** Hard ceilings: rows are data rows per sheet; cells and text are totals including headers. */
export interface ExportLimits {
  readonly maxRows?: number;
  readonly maxCells?: number;
  readonly maxTextCharacters?: number;
  readonly maxOutputBytes?: number;
}

/** @internal Check before accepting the next value or chunk. */
export class ExportBudget {
  private cells = 0;
  private text = 0;
  private reservedCells = 0;
  private reservedText = 0;
  readonly limits: Readonly<ExportLimits>;
  constructor(limits: ExportLimits = {}) {
    for (const [key, value] of Object.entries(limits))
      if (value !== undefined && (!Number.isSafeInteger(value) || value < 0)) throw new RangeError(key + " must be a nonnegative safe integer.");
    this.limits = Object.freeze({ ...limits });
  }
  check(kind: keyof ExportLimits, value: number): void {
    const maximum = this.limits[kind];
    if (maximum !== undefined && value > maximum) throw new OfficeIMOError("RESOURCE_LIMIT", kind + " exceeded (" + value + " > " + maximum + ").");
  }
  row(count: number): void { this.check("maxRows", count); }
  reserve(cells: number, characters: number): void {
    this.check("maxCells", this.cells + this.reservedCells + cells);
    this.check("maxTextCharacters", this.text + this.reservedText + characters);
    this.reservedCells += cells; this.reservedText += characters;
  }
  release(cells: number, characters: number): void { this.reservedCells -= cells; this.reservedText -= characters; }
  cell(value: unknown, reservedCharacters?: number): void {
    if (reservedCharacters !== undefined) this.release(1, reservedCharacters);
    // Immutable limits cannot acquire a ceiling later. Avoid per-cell accounting
    // when neither emitted-cell nor text totals are observable by a resource cap.
    if (this.limits.maxCells === undefined && this.limits.maxTextCharacters === undefined) return;
    this.check("maxCells", this.cells + this.reservedCells + 1);
    const length = typeof value === "string" ? value.length : 0;
    this.check("maxTextCharacters", this.text + this.reservedText + length);
    this.cells++; this.text += length;
  }
}

/** @internal Ownership stays with the caller. */
export function boundedSink(sink: ByteSink, budget: ExportBudget): ByteSink {
  let bytes = 0;
  return { write(chunk) {
    budget.check("maxOutputBytes", bytes + chunk.length);
    bytes += chunk.length;
    return sink.write(chunk);
  } };
}
