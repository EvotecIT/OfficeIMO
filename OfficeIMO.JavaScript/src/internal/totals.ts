import type { ProjectionColumn } from "./columns.js";

/** Shared numeric report totals. Count follows spreadsheet COUNT and ignores nonnumeric cells. */
export type TotalOperation = "sum" | "count" | "average" | "min" | "max";
export class NumericAggregate {
  count = 0;
  private sum = 0;
  private mean = 0;
  private min = Infinity;
  private max = -Infinity;
  constructor(readonly operation: TotalOperation | undefined) {}
  accept(value: unknown): void {
    if (!this.operation || typeof value !== "number" || !Number.isFinite(value)) return;
    this.count++;
    if (this.operation === "sum") {
      this.sum += value;
      if (!Number.isFinite(this.sum)) throw new RangeError("Numeric total exceeds finite number range.");
    } else if (this.operation === "average") {
      const delta = value - this.mean;
      this.mean = this.count === 1 ? value : Number.isFinite(delta) ? this.mean + delta / this.count : this.mean * ((this.count - 1) / this.count) + value / this.count;
    } else if (this.operation === "min") this.min = Math.min(this.min, value);
    else if (this.operation === "max") this.max = Math.max(this.max, value);
  }
  value(): number | null {
    return this.operation === "count" ? this.count : this.operation === "sum" ? this.sum : !this.count ? null : this.operation === "average" ? this.mean : this.operation === "min" ? this.min : this.max;
  }
}
export function createTotals(columns: readonly ProjectionColumn[], totals: Readonly<Record<string, TotalOperation>> = {}): NumericAggregate[] {
  const keys = columns.map(c => c.key ?? c.header);
  for (const [key, operation] of Object.entries(totals)) {
    if (keys.filter(k => k === key).length !== 1) throw new TypeError("Totals need an unambiguous declared column key: " + key);
    if (!["sum", "count", "average", "min", "max"].includes(operation)) throw new TypeError("Invalid total operation.");
  }
  return keys.map(key => new NumericAggregate(Object.prototype.hasOwnProperty.call(totals, key) ? totals[key] : undefined));
}
