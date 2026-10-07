import type { Border, Font } from "./styles.js";

/** A1 cell/range, or data-only columns resolved after streaming finishes. Column numbers are one-based; strings are declared column keys. */
export type ConditionalRange = string | { readonly column: string | number; readonly through?: string | number };
/** Only supplied properties override the cell. Omitting numberFormat preserves the existing number/date format. */
export interface ConditionalStyle {
  readonly font?: Pick<Font, "bold" | "italic" | "underline" | "strike" | "color">;
  readonly fill?: { readonly color: string };
  readonly border?: Border;
  readonly numberFormat?: string;
}
export type ConditionalOperator = "equal" | "notEqual" | "greaterThan" | "greaterThanOrEqual" | "lessThan" | "lessThanOrEqual";
/** Formula text uses Excel's invariant syntax. It is stored, never evaluated by this writer. */
export type ConditionalThreshold = { readonly type: "min" | "max" } |
  { readonly type: "number" | "percent" | "percentile"; readonly value: number } |
  { readonly type: "formula"; readonly value: string };
export interface ConditionalColorStop { readonly threshold: ConditionalThreshold; readonly color: string; }
interface ConditionalBase { readonly range: ConditionalRange; }
interface ConditionalHighlight extends ConditionalBase { readonly style: ConditionalStyle; readonly stopIfTrue?: boolean; }
/** Array order sets worksheet-wide priority, starting at one. Overlapping ranges are allowed. */
export type ConditionalFormat =
  | (ConditionalHighlight & { readonly type: "cellIs"; readonly operator: ConditionalOperator; readonly value: number })
  | (ConditionalHighlight & { readonly type: "cellIs"; readonly operator: "between" | "notBetween"; readonly values: readonly [number, number] })
  | (ConditionalHighlight & { readonly type: "expression"; readonly formula: string })
  | (ConditionalBase & { readonly type: "colorScale"; readonly stops: readonly [ConditionalColorStop, ConditionalColorStop] | readonly [ConditionalColorStop, ConditionalColorStop, ConditionalColorStop] })
  | (ConditionalBase & { readonly type: "dataBar"; readonly color: string; readonly minimum?: ConditionalThreshold; readonly maximum?: ConditionalThreshold; readonly showValue?: boolean });
