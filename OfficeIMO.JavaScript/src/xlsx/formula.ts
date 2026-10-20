import { NotSupportedError } from "../core/errors.js";
import { cleanXml, escapeXml } from "../xml/index.js";

const formulaCells = new WeakSet<object>();

/** @internal Reject direct external-workbook and DDE references without interpreting Excel expressions. */
function localReferences(formula: string): void {
  for (let i = 0; i < formula.length; i++) {
    const character = formula[i];
    if (character === '"') {
      // Excel string literals double their embedded quotes. Their text is not a reference token.
      while (++i < formula.length) if (formula[i] === '"') {
        if (formula[i + 1] !== '"') break;
        i++;
      }
    } else if (character === "'") {
      let quoted = "";
      while (++i < formula.length) {
        if (formula[i] === "'") {
          if (formula[i + 1] !== "'") break;
          i++;
        }
        quoted += formula[i];
      }
      if (formula.slice(i + 1).trimStart().startsWith("!") && /[\[\]]/.test(quoted))
        throw new NotSupportedError("external-workbook formulas");
    } else if (character === "[") {
      let depth = 1;
      while (++i < formula.length && depth) {
        // Apostrophes escape special characters in table-column captions.
        if (formula[i] === "'") { i++; continue; }
        if (formula[i] === "[") depth++;
        else if (formula[i] === "]") depth--;
      }
      if (/^[A-Za-z0-9_.]*!/.test(formula.slice(i)))
        throw new NotSupportedError("external-workbook formulas");
      i--;
    } else if (character === "|") throw new NotSupportedError("DDE formulas");
  }
}

/** @internal Excel invariant expression text, stored without its optional leading equals sign. */
export function formulaText(value: string): string {
  if (typeof value !== "string") throw new TypeError("Formula text must be a string.");
  const text = value.trim().replace(/^=/, "");
  if (!text.trim() || text.length > 8192) throw new RangeError("Formulas require 1 through 8,192 characters.");
  // Text stripping could change a calculation, so formulas always reject invalid XML characters.
  cleanXml(text, "reject");
  localReferences(text);
  return text;
}

/** An authored Excel formula and its initial numeric result. The writer stores both without evaluating the expression. */
export class FormulaCell {
  static [Symbol.hasInstance](value: unknown): boolean {
    return value !== null && typeof value === "object" && formulaCells.has(value);
  }
  readonly formula: string;
  constructor(formula: string, readonly cachedValue: number, readonly style?: number) {
    this.formula = formulaText(formula);
    if (typeof cachedValue !== "number" || !Number.isFinite(cachedValue)) throw new TypeError("Formula caches must be finite numbers.");
    if (style !== undefined && (!Number.isSafeInteger(style) || style < 0)) throw new RangeError("Formula style indexes must be nonnegative safe integers.");
    formulaCells.add(this);
    Object.freeze(this);
  }
}

/** @internal The worksheet owns the cell coordinate, style and cumulative formula budget. */
export function formulaXml(cell: FormulaCell): string {
  if (!(cell instanceof FormulaCell)) throw new TypeError("Formula cells must be created with FormulaCell.");
  return "<f>" + escapeXml(cell.formula, "reject") + "</f><v>" + cell.cachedValue + "</v>";
}
