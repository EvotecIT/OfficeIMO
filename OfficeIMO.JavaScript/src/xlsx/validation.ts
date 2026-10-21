import { OfficeIMOError } from "../core/errors.js";
import { checkAbort } from "../core/iteration.js";
import type { ProjectionColumn } from "../internal/columns.js";
import { cleanXml, escapeOoxmlAttribute, escapeXml } from "../xml/index.js";
import { fields, optionalBoolean } from "./differential-styles.js";
import { resolveDataColumn } from "./ranges.js";
import type { ListValidation } from "./validation-types.js";
import { dataRowReference } from "./data-row-ranges.js";
import type { DataRowRange } from "./data-row-ranges.js";

/** @internal Captured worksheet metadata; no caller arrays or option objects remain. */
export interface PreparedValidation { readonly column: number; readonly attributes: string; readonly formula: string; }
/** @internal Counts are charged cumulatively by the workbook only after successful registration. */
export interface PreparedValidations { readonly rules: readonly PreparedValidation[]; readonly count: number; readonly textCharacters: number; }

function message(value: string | undefined, maximum: number, name: string): string | undefined {
  if (value === undefined) return undefined;
  if (typeof value !== "string") throw new TypeError(name + " must be text.");
  if (value.length > maximum) throw new RangeError(name + " exceeds " + maximum + " UTF-16 units.");
  return cleanXml(value, "reject");
}

/** @internal Snapshot and validate before any source iteration or workbook registration. */
export function prepareDataValidations(rules: readonly ListValidation[] | undefined, columns: readonly ProjectionColumn[], availableCount = 1000,
  availableText = 1000000, signal?: AbortSignal): PreparedValidations {
  checkAbort(signal);
  if (rules === undefined) return Object.freeze({ rules: Object.freeze([]), count: 0, textCharacters: 0 });
  if (!Array.isArray(rules)) throw new TypeError("dataValidation must be an array.");
  const count = rules.length;
  if (count > availableCount) throw new OfficeIMOError("RESOURCE_LIMIT", "maxDataValidations exceeded.");
  const prepared: PreparedValidation[] = [];
  const usedColumns = new Set<number>();
  let textCharacters = 0;
  const acceptText = (length: number) => {
    if (textCharacters + length > availableText) throw new OfficeIMOError("RESOURCE_LIMIT", "maxValidationTextCharacters exceeded.");
    textCharacters += length;
  };
  for (let index = 0; index < count; index++) {
    checkAbort(signal);
    const rule = rules[index]!;
    fields(rule, ["type", "column", "items", "allowBlank", "showErrorMessage", "errorStyle", "errorTitle", "errorMessage"], "data validation");
    const { type, column, items, allowBlank, showErrorMessage, errorStyle, errorTitle, errorMessage } = rule;
    if (type !== "list") throw new TypeError("Only inline-list data validation is supported.");
    const selected = resolveDataColumn(columns, column, "Validation");
    if (usedColumns.has(selected)) throw new TypeError("A data column cannot have multiple validation rules.");
    usedColumns.add(selected);
    optionalBoolean(allowBlank, "allowBlank"); optionalBoolean(showErrorMessage, "showErrorMessage");
    if (errorStyle !== undefined && !["stop", "warning", "information"].includes(errorStyle)) throw new TypeError("Unsupported validation error style.");
    if (!Array.isArray(items)) throw new TypeError("Inline validation lists require a nonempty string array.");
    // Even empty choices need separators. Bound the loop before observing any item getters.
    const itemCount = items.length;
    if (!itemCount) throw new TypeError("Inline validation lists require a nonempty string array.");
    if (itemCount > 254) throw new RangeError("Inline validation lists must fit 255 UTF-16 units including quotes and separators.");
    let formula = '"';
    for (let item = 0; item < itemCount; item++) {
      checkAbort(signal);
      const value = items[item];
      if (typeof value !== "string") throw new TypeError("Inline validation choices must be strings.");
      if (value.length > 253 || value.includes(",")) throw new RangeError("Inline validation choices cannot contain commas and must fit the 255-character list source.");
      cleanXml(value, "reject");
      const escaped = value.replace(/"/g, '""');
      if (formula.length + (item ? 1 : 0) + escaped.length + 1 > 255) throw new RangeError("Inline validation lists must fit 255 UTF-16 units including quotes and separators.");
      formula += (item ? "," : "") + escaped;
    }
    formula += '"';
    const title = message(errorTitle, 32, "Validation error title"), error = message(errorMessage, 255, "Validation error message");
    acceptText(formula.length + (title?.length ?? 0) + (error?.length ?? 0));
    const attributes = ' type="list" allowBlank="' + (allowBlank === false ? 0 : 1) + '" showErrorMessage="' + (showErrorMessage === false ? 0 : 1) +
      '" errorStyle="' + (errorStyle ?? "stop") + '"' + (title === undefined ? "" : ' errorTitle="' + escapeOoxmlAttribute(title, "reject") + '"') +
      (error === undefined ? "" : ' error="' + escapeOoxmlAttribute(error, "reject") + '"');
    prepared.push(Object.freeze({ column: selected, attributes, formula }));
  }
  checkAbort(signal);
  return Object.freeze({ rules: Object.freeze(prepared), count, textCharacters });
}

/** @internal Empty sources do not create validation over headings, totals or an artificial blank row. */
export function* dataValidationsXml(prepared: PreparedValidations, headerRows: number, dataRows: number, ranges?: readonly DataRowRange[]): Iterable<string> {
  if (!prepared.count || !dataRows || ranges?.length === 0) return;
  yield '<dataValidations count="' + prepared.count + '">';
  for (const rule of prepared.rules) {
    const reference = dataRowReference(rule.column, rule.column, headerRows, dataRows, ranges)!;
    yield '<dataValidation' + rule.attributes + ' sqref="' + reference + '"><formula1>' + escapeXml(rule.formula, "reject") + '</formula1></dataValidation>';
  }
  yield '</dataValidations>';
}
