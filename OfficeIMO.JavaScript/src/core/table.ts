import type { ExportValue } from "./presentation.js";

/** A heading/footer anchor. Covered matrix positions are null; rows have the declared column count. */
export interface TableSpanCell {
  readonly value: ExportValue;
  readonly columnSpan?: number;
  readonly rowSpan?: number;
}
export type TableSpanRows = readonly (readonly (TableSpanCell | null)[])[];
