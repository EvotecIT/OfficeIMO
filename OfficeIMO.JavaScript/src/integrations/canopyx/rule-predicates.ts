import { columnName, excelDate } from "../../xlsx/values.js";
import { captureFields } from "../../internal/capture.js";
import { datetime } from "./dates.js";
import type { CanopyCapture, CanopyRuleFilter, CanopyRulePredicate, CanopySourceOptions } from "./types.js";

/** Compile only comparisons whose raw types and native encodings have a faithful bounded mapping. */
export function canopyPredicate(request: CanopyCapture["request"], grouped: boolean, options: CanopySourceOptions,
  root: CanopyRulePredicate, unsupported: (code: string) => never, cellTypes: Map<number, "number" | "date">): string {
  let nodes = 0;
  const operations: Readonly<Record<string, string>> = { eq: "=", ne: "<>", gt: ">", ge: ">=", lt: "<", le: "<=" };
  const numeric = (value: string | number | boolean): number => {
    const number = typeof value === "boolean" ? Number(value) : typeof value === "string" ? Number(value) : value;
    // Use platform decimal conversion; separator/culture profiles need a diagnostic, not a second parser.
    if (typeof value === "string") {
      const text = value.trim().toLowerCase();
      if (Number.parseFloat(value) !== number || ["0x", "0b", "0o"].some(prefix => text.startsWith(prefix)))
        return unsupported("UNSUPPORTED_NUMERIC_RULE_LITERAL");
    }
    if (!Number.isFinite(number) || Math.abs(number) > Number.MAX_SAFE_INTEGER) throw new TypeError("Canopy rule comparison values must be finite safe-range numbers.");
    return number;
  };
  const date = (value: string | number | boolean): number => {
    if (options.datetime !== "typed" || (options.xlsx?.dateMode ?? request.timeZone) !== "utc")
      return unsupported("DATETIME_RULE_REQUIRES_TYPED_UTC");
    if (typeof value === "boolean") return unsupported("UNSUPPORTED_DATETIME_RULE_LITERAL");
    if (typeof value === "number" && !Number.isSafeInteger(value)) return unsupported("UNSUPPORTED_DATETIME_RULE_LITERAL");
    try {
      const literal = typeof value === "number" ? new Date(value).toISOString() : value;
      const captured = datetime(literal, "typed", "utc");
      if (!(captured instanceof Date)) return unsupported("UNSUPPORTED_DATETIME_RULE_LITERAL");
      const serial = excelDate(captured, "utc");
      if (serial === null) return unsupported("UNSUPPORTED_DATETIME_RULE_LITERAL");
      return serial;
    } catch (error) {
      if (!(error instanceof TypeError) && !(error instanceof RangeError)) throw error;
      return unsupported("UNSUPPORTED_DATETIME_RULE_LITERAL");
    }
  };
  const leaf = (input: CanopyRuleFilter): string => {
    if (++nodes > 64) return unsupported("RULE_PREDICATE_LIMIT");
    const filter = captureFields(input, ["column", "operator", "values"]), operation = operations[filter.operator];
    if (Object.keys(filter).some(key => !["column", "operator", "values"].includes(key))) return unsupported("UNSUPPORTED_RULE_FILTER_OPTIONS");
    const index = request.columns.findIndex(column => column.id === filter.column);
    if (index < 0) return unsupported("RULE_COLUMN_NOT_EXPORTED:" + filter.column);
    const column = request.columns[index]!, letter = columnName(index + 1 + (grouped ? 1 : 0));
    const reference = "INDEX($" + letter + ":$" + letter + ",ROW())", present = "NOT(ISBLANK(" + reference + "))";
    if (column.kind === "number") cellTypes.set(index, "number");
    if (column.kind === "datetime") {
      if (options.datetime !== "typed" || (options.xlsx?.dateMode ?? request.timeZone) !== "utc")
        return unsupported("DATETIME_RULE_REQUIRES_TYPED_UTC");
      cellTypes.set(index, "date");
    }
    if (filter.operator === "isNull" || filter.operator === "isNotNull") {
      if (filter.values !== undefined && filter.values.length) return unsupported("UNSUPPORTED_NULL_RULE_OPERANDS");
      return filter.operator === "isNull" ? "ISBLANK(" + reference + ")" : present;
    }
    if (!operation || !Array.isArray(filter.values) || filter.values.length !== 1) return unsupported("UNSUPPORTED_RULE_FILTER:" + filter.operator);
    const value = filter.values[0];
    if (typeof value !== "string" && typeof value !== "number" && typeof value !== "boolean") throw new TypeError("Canopy rule values must be portable scalars.");
    if (column.kind === "number" || column.kind === "datetime") {
      const operand = column.kind === "number" ? numeric(value) : date(value);
      return "AND(" + present + ",ISNUMBER(" + reference + ")," + reference + operation + String(operand) + ")";
    }
    if (filter.operator !== "eq" && filter.operator !== "ne") return unsupported("UNSUPPORTED_TEXT_RULE_COMPARISON");
    // Canopy's Unicode text folding is not the native workbook's comparison contract.
    if (typeof value === "string") return unsupported("UNQUALIFIED_NATIVE_TEXT_COMPARISON");
    const kind = typeof value === "boolean" ? "ISLOGICAL" : "ISNUMBER";
    const literal = typeof value === "boolean" ? value ? "TRUE" : "FALSE" : String(numeric(value));
    const sameType = kind + "(" + reference + ")", comparison = reference + operation + literal;
    return filter.operator === "eq" ? "AND(" + present + "," + sameType + "," + comparison + ")"
      : "AND(" + present + ",OR(NOT(" + sameType + ")," + comparison + "))";
  };
  const visit = (input: CanopyRulePredicate, depth: number): string => {
    if (depth > 4 || ++nodes > 64) return unsupported("RULE_PREDICATE_LIMIT");
    if (!input || typeof input !== "object") throw new TypeError("Canopy rule predicate must be an object.");
    const captured = captureFields(input, ["operator", "filters", "groups"]);
    if (captured.operator !== "and" && captured.operator !== "or") return unsupported("UNSUPPORTED_RULE_BOOLEAN:" + captured.operator);
    const filters: readonly CanopyRuleFilter[] = captured.filters ?? [], groups: readonly CanopyRulePredicate[] = captured.groups ?? [];
    if (!Array.isArray(filters) || !Array.isArray(groups)) throw new TypeError("Canopy predicate filters and groups must be arrays.");
    const terms: string[] = [];
    for (const filter of filters) terms.push(leaf(filter));
    for (const group of groups) terms.push(visit(group, depth + 1));
    if (!terms.length) return unsupported("EMPTY_RULE_PREDICATE");
    return captured.operator.toUpperCase() + "(" + terms.join(",") + ")";
  };
  const formula = visit(root, 1);
  if (formula.length > 8192) return unsupported("RULE_FORMULA_LIMIT");
  return formula;
}
