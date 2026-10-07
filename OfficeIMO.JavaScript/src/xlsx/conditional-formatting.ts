import type { Column } from "../core/index.js";
import { cleanXml, escapeXml, escapeOoxmlAttribute } from "../xml/index.js";
import type { InvalidCharacterPolicy } from "../xml/index.js";
import { cellPosition } from "./attachments.js";
import { columnName } from "./values.js";
import { colorArgb } from "./style-xml.js";
import { fields, optionalBoolean, differentialXml } from "./differential-styles.js";
import type { ConditionalFormat, ConditionalRange, ConditionalStyle, ConditionalThreshold } from "./conditional-types.js";
import type { StyleRegistry } from "./styles.js";

interface Target { readonly first: number; readonly last: number; readonly reference?: string; readonly bottom?: number; }
/** @internal Validated metadata, without workbook registrations. */
export interface PreparedRule { readonly target: Target; readonly type: string; readonly attributes: string; readonly body: string; readonly style?: ConditionalStyle; }
function target(range: ConditionalRange, columns: readonly Column[]): Target {
  if (typeof range === "string") {
    const parts = range.split(":");
    if (parts.length > 2) throw new TypeError("Conditional ranges require one uppercase A1 cell or rectangle.");
    const first = cellPosition(parts[0]!), last = cellPosition(parts[1] ?? parts[0]!);
    if (last.row < first.row || last.column < first.column || last.column > columns.length) throw new RangeError("Conditional ranges require ordered cells within declared columns.");
    return { first: first.column, last: last.column, reference: range, bottom: last.row };
  }
  fields(range, ["column", "through"], "conditional data range");
  function resolve(value: string | number): number {
    if (typeof value === "string") {
      const matches = columns.flatMap((column, i) => column.key === value ? [i + 1] : []);
      if (matches.length !== 1) throw new TypeError("Conditional column keys must identify exactly one declared column.");
      return matches[0]!;
    }
    if (!Number.isInteger(value) || value < 1 || value > columns.length) throw new RangeError("Conditional column numbers must be one-based and within declared columns.");
    return value;
  }
  const first = resolve(range.column), last = range.through === undefined ? first : resolve(range.through);
  if (last < first) throw new RangeError("Conditional data ranges require ordered columns.");
  return { first, last };
}
function finite(value: number): string {
  if (typeof value !== "number" || !Number.isFinite(value)) throw new TypeError("Conditional values must be finite numbers.");
  return String(value);
}
function formula(value: string): string {
  if (typeof value !== "string") throw new TypeError("Conditional formulas must be strings.");
  const text = value.trim().replace(/^=/, "");
  if (!text.trim() || text.length > 8192) throw new RangeError("Conditional formulas require 1 through 8,192 characters.");
  // Stripping invalid characters could change a calculation; reject them even under the text-strip policy.
  cleanXml(text, "reject"); return text;
}
function threshold(value: ConditionalThreshold, position: "first" | "middle" | "last"): string {
  fields(value, value.type === "min" || value.type === "max" ? ["type"] : ["type", "value"], "conditional threshold");
  if (value.type === "min" || value.type === "max") {
    if (value.type !== (position === "first" ? "min" : position === "last" ? "max" : "")) throw new TypeError("Use min only at the start and max only at the end of a scale.");
    return '<cfvo type="' + value.type + '"/>';
  }
  if (value.type === "formula") return '<cfvo type="formula" val="' + escapeOoxmlAttribute(formula(value.value), "reject") + '"/>';
  if (!["number", "percent", "percentile"].includes(value.type)) throw new TypeError("Unsupported conditional threshold type.");
  if (!("value" in value)) throw new TypeError("A numeric threshold requires a value.");
  const number = finite(value.value);
  if (value.type !== "number" && (value.value < 0 || value.value > 100)) throw new RangeError("Conditional percent/percentile thresholds must be from 0 through 100.");
  return '<cfvo type="' + (value.type === "number" ? "num" : value.type) + '" val="' + number + '"/>';
}
function ordered(values: readonly ConditionalThreshold[]): void {
  for (let i = 1; i < values.length; i++) {
    const before = values[i - 1]!, after = values[i]!;
    if ("value" in before && "value" in after && before.type === after.type && typeof before.value === "number" && typeof after.value === "number" && before.value > after.value)
      throw new RangeError("Conditional thresholds of the same numeric type must be ordered.");
  }
}
/** @internal Compile bounded metadata once, outside the row hot path. */
export function prepareConditionalFormats(rules: readonly ConditionalFormat[] | undefined, columns: readonly Column[], policy: InvalidCharacterPolicy): readonly PreparedRule[] {
  if (rules === undefined) return [];
  if (!Array.isArray(rules)) throw new TypeError("conditionalFormats must be an array.");
  return (rules as readonly ConditionalFormat[]).map(rule => {
    const common = ["type", "range"], highlighting = ["style", "stopIfTrue"];
    fields(rule, [...common, ...(rule?.type === "cellIs" ? [...highlighting, "operator", "value", "values"] : rule?.type === "expression" ? [...highlighting, "formula"] : rule?.type === "colorScale" ? ["stops"] : rule?.type === "dataBar" ? ["color", "minimum", "maximum", "showValue"] : [])], "conditional rule");
    const selected = target(rule.range, columns);
    let attributes = "", body = "", style: ConditionalStyle | undefined;
    if (rule.type === "cellIs" || rule.type === "expression") {
      optionalBoolean(rule.stopIfTrue, "stopIfTrue");
      differentialXml(rule.style, policy, () => 164);
      style = { ...rule.style, ...(rule.style.font ? { font: { ...rule.style.font } } : {}), ...(rule.style.fill ? { fill: { ...rule.style.fill } } : {}),
        ...(rule.style.border ? { border: Object.fromEntries(Object.entries(rule.style.border).filter(([, edge]) => edge !== undefined).map(([side, edge]) => [side, { ...edge }])) } : {}) };
      if (rule.stopIfTrue !== undefined) attributes += ' stopIfTrue="' + (rule.stopIfTrue ? 1 : 0) + '"';
      if (rule.type === "expression") body = '<formula>' + escapeXml(formula(rule.formula), "reject") + '</formula>';
      else {
        const pair = rule.operator === "between" || rule.operator === "notBetween";
        if (!["equal", "notEqual", "greaterThan", "greaterThanOrEqual", "lessThan", "lessThanOrEqual", "between", "notBetween"].includes(rule.operator)) throw new TypeError("Unsupported conditional comparison operator.");
        if (pair) {
          if ("value" in rule || !Array.isArray(rule.values) || rule.values.length !== 2) throw new TypeError("between/notBetween require exactly two values.");
          body = rule.values.map(value => '<formula>' + finite(value) + '</formula>').join("");
          if (rule.values[0] > rule.values[1]) throw new RangeError("Conditional comparison bounds must be ordered.");
        } else {
          if ("values" in rule || !("value" in rule)) throw new TypeError("A conditional comparison requires one value.");
          body = '<formula>' + finite(rule.value) + '</formula>';
        }
        attributes += ' operator="' + rule.operator + '"';
      }
    } else if (rule.type === "colorScale") {
      if (!Array.isArray(rule.stops) || (rule.stops.length !== 2 && rule.stops.length !== 3)) throw new TypeError("Color scales require two or three stops.");
      const colors: string[] = [], thresholds: ConditionalThreshold[] = [];
      body = '<colorScale>' + rule.stops.map((stop, i) => {
        fields(stop, ["threshold", "color"], "color stop"); colors.push(colorArgb(stop.color)); thresholds.push(stop.threshold);
        return threshold(stop.threshold, i === 0 ? "first" : i === rule.stops.length - 1 ? "last" : "middle");
      }).join("") + colors.map(color => '<color rgb="' + color + '"/>').join("") + '</colorScale>';
      ordered(thresholds);
    } else if (rule.type === "dataBar") {
      optionalBoolean(rule.showValue, "showValue");
      const lower = rule.minimum === undefined ? { type: "min" as const } : rule.minimum, upper = rule.maximum === undefined ? { type: "max" as const } : rule.maximum;
      body = '<dataBar showValue="' + (rule.showValue === false ? 0 : 1) + '">' + threshold(lower, "first") + threshold(upper, "last") + '<color rgb="' + colorArgb(rule.color) + '"/></dataBar>';
      ordered([lower, upper]);
    } else throw new TypeError("Unsupported conditional rule type.");
    return { target: selected, type: rule.type, attributes, body, ...(style ? { style } : {}) };
  });
}
/** @internal Immutable worksheet-owned rule metadata; no cells or rows are retained. */
export class ConditionalFormats {
  private readonly rules: readonly { target: Target; xml: string }[];
  constructor(prepared: readonly PreparedRule[], styles: StyleRegistry) {
    const differential = styles.addDifferentials(prepared.flatMap(rule => rule.style ? [rule.style] : []));
    let styleIndex = 0;
    this.rules = prepared.map((rule, i) => ({ target: rule.target,
      xml: '<cfRule type="' + rule.type + '" priority="' + (i + 1) + '"' + (rule.style ? ' dxfId="' + differential[styleIndex++] + '"' : "") + rule.attributes + '>' + rule.body + '</cfRule>' }));
  }
  get count(): number { return this.rules.length; }
  *xml(headerRows: number, dataRows: number, totalRows: number): Iterable<string> {
    for (const rule of this.rules) {
      if (rule.target.bottom !== undefined && rule.target.bottom > totalRows) throw new RangeError("Conditional ranges must stay within exported rows.");
      if (!rule.target.reference && !dataRows) continue;
      const reference = rule.target.reference ?? columnName(rule.target.first) + (headerRows + 1) + ':' + columnName(rule.target.last) + (headerRows + dataRows);
      yield '<conditionalFormatting sqref="' + reference + '">' + rule.xml + '</conditionalFormatting>';
    }
  }
}
