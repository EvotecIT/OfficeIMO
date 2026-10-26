import type { CellPresentation } from "../../core/index.js";
import type { ConditionalFormat, ConditionalStyle, ConditionalThreshold, ConditionalColorStop } from "../../xlsx/index.js";
import { captureFields } from "../../internal/capture.js";
import { canopyPredicate } from "./rule-predicates.js";
import type { CanopyCapture, CanopyDiagnostic, CanopyFormat, CanopyFormatRule, CanopySourceOptions } from "./types.js";

// Stable document palette. Browser CSS variables and resolved cell.format are never interpreted.
const palette: Readonly<Record<string, { readonly color: string; readonly background: string; readonly foreground: string }>> = {
  info: { color: "4472C4", background: "DDEBF7", foreground: "1F4E78" },
  success: { color: "70AD47", background: "E2F0D9", foreground: "375623" },
  warning: { color: "FFC000", background: "FFF2CC", foreground: "9C6500" },
  danger: { color: "C00000", background: "FFC7CE", foreground: "9C0006" }
};
class UnsupportedRule extends Error {}
const unsupported = (message: string): never => { throw new UnsupportedRule(message); };

/** Compile declarations in source order, then emit native target/aspect priorities.
 * Data-column selectors let grouped writers exclude generated headings, subtotals and grand totals. */
export function canopyRules(request: CanopyCapture["request"], format: CanopyFormat, grouped: boolean,
  tones: ReadonlyMap<string, Readonly<CellPresentation>>, report: (diagnostic: CanopyDiagnostic) => void,
  options: CanopySourceOptions = {}, cellTypes?: Map<number, "number" | "date">): readonly ConditionalFormat[] {
  if (request.formatRules === undefined || format === "csv") return [];
  if (!Array.isArray(request.formatRules) || request.formatRules.length > 1000) throw new TypeError("Canopy formatRules must be a bounded array of at most 1,000 declarations.");
  const visuals: ConditionalFormat[] = [], cells: ConditionalFormat[] = [], rows: ConditionalFormat[] = [], ids = new Set<string>(), columns = request.columns;
  const inputs: readonly CanopyFormatRule[] = request.formatRules;
  const declarations = inputs.map(input => {
    if (!input || typeof input !== "object") throw new TypeError("Canopy rule declarations must be objects.");
    const rule = captureFields(input, ["id", "enabled", "stopIfTrue", "column", "when", "target", "style"]);
    if (typeof rule.id !== "string" || !rule.id || ids.has(rule.id)) throw new TypeError("Canopy rule IDs must be unique nonempty strings.");
    ids.add(rule.id);
    if (rule.enabled !== undefined && typeof rule.enabled !== "boolean" || rule.stopIfTrue !== undefined && typeof rule.stopIfTrue !== "boolean") throw new TypeError("Canopy rule flags must be boolean.");
    const style = rule.style && typeof rule.style === "object" ? captureFields(rule.style, ["tone", "bar", "scale"]) : rule.style;
    return { ...rule, style };
  });
  if (format === "pdf") {
    for (const rule of declarations) if (rule.enabled !== false) report({ code: "PDF_NATIVE_RULE_NOT_RENDERED:" + rule.id });
    return [];
  }
  const hasStops = declarations.some(rule => rule.enabled !== false && rule.stopIfTrue);
  if (hasStops && declarations.some(rule => rule.enabled !== false && (rule.style?.bar !== undefined || rule.style?.scale !== undefined))) {
    report({ code: "UNSUPPORTED_GLOBAL_STOP_WITH_NATIVE_VISUALS" }); return [];
  }
  const barColumns = new Set(declarations.filter(rule => rule.enabled !== false && rule.style?.bar !== undefined).map(rule => rule.column));
  const barTone = declarations.find(rule => rule.enabled !== false && rule.target !== "row" && rule.style?.tone !== undefined && rule.column !== undefined && barColumns.has(rule.column));
  if (barTone) {
    // Canopy suppresses a cell tone's fill only where its bar resolves. A single native tone patch cannot express both cases.
    report({ code: "UNSUPPORTED_CELL_TONE_WITH_DATA_BAR", columnId: barTone.column! }); return [];
  }
  const priorStops: string[] = [];
  const priorRows: string[] = [], priorCells = new Map<string, string[]>();
  const requiredCells = new Map<number, "number" | "date">();
  const column = (id: string): number => {
    const index = columns.findIndex(column => column.id === id);
    if (index < 0) return unsupported("RULE_COLUMN_NOT_EXPORTED:" + id);
    return index;
  };
  const numeric = (id: string): void => {
    if (columns[column(id)]!.kind !== "number") unsupported("RULE_COLUMN_NOT_NUMERIC:" + id);
  };
  const color = (name: string): string => {
    const mapping = tones.get(name), value = mapping?.background ?? mapping?.color ?? palette[name]?.color;
    if (value === undefined) return unsupported("UNMAPPED_RULE_TONE:" + name);
    return value;
  };
  const style = (name: string): ConditionalStyle => {
    const defaults = palette[name];
    const mapping: Readonly<CellPresentation> | undefined = tones.get(name) ?? (defaults ? { background: defaults.background, color: defaults.foreground } : undefined);
    if (!mapping) return unsupported("UNMAPPED_RULE_TONE:" + name);
    if ("alignment" in mapping || "wrapText" in mapping) return unsupported("UNSUPPORTED_RULE_TONE_STYLE:" + name);
    return {
      ...("background" in mapping && mapping.background !== undefined ? { fill: { color: mapping.background } } : {}),
      font: { ...(mapping.color === undefined ? {} : { color: mapping.color }),
        ...("bold" in mapping && mapping.bold !== undefined ? { bold: mapping.bold } : {}),
        ...("italic" in mapping && mapping.italic !== undefined ? { italic: mapping.italic } : {}) },
      ...("numberFormat" in mapping && mapping.numberFormat !== undefined ? { numberFormat: mapping.numberFormat } : {})
    };
  };
  for (const rule of declarations) {
    if (rule.enabled === false) continue;
    try {
      if (request.values !== "raw") unsupported("NATIVE_RULE_REQUIRES_RAW_VALUES");
      if (rule.target === "columns" || "columns" in rule && rule.columns !== undefined) unsupported("UNSUPPORTED_RULE_COLUMNS_TARGET");
      if (!rule.style || typeof rule.style !== "object" || Array.isArray(rule.style)) throw new TypeError("Canopy rule style must be an object.");
      const styleInput = captureFields(rule.style, ["tone", "bar", "scale"]);
      if (Object.keys(styleInput).some(key => !["tone", "bar", "scale"].includes(key))) unsupported("UNSUPPORTED_RULE_STYLE");
      const supplied = [styleInput.bar, styleInput.scale, styleInput.tone].filter(value => value !== undefined);
      if (supplied.length !== 1) unsupported("UNSUPPORTED_RULE_STYLE_COMBINATION");
      if (styleInput.bar !== undefined) {
        if (rule.column === undefined || rule.when !== undefined || rule.target === "row") unsupported("UNSUPPORTED_BAR_TARGET");
        numeric(rule.column!);
        if (rule.stopIfTrue) unsupported("UNSUPPORTED_BAR_STOP");
        const { minimum, maximum, tone = "info", showValue } = captureFields(styleInput.bar, ["minimum", "maximum", "tone", "showValue"]);
        if (minimum === undefined || maximum === undefined) return unsupported("UNBOUNDED_CANOPY_BAR");
        if (!Number.isFinite(minimum) || !Number.isFinite(maximum) || maximum <= minimum) throw new TypeError("Canopy bar bounds must be finite and increasing.");
        visuals.push({ type: "dataBar", range: { column: rule.column! }, color: color(tone),
          minimum: { type: "number", value: minimum }, maximum: { type: "number", value: maximum }, ...(showValue === undefined ? {} : { showValue }) });
        requiredCells.set(column(rule.column!), "number");
      } else if (styleInput.scale !== undefined) {
        if (rule.stopIfTrue) unsupported("UNSUPPORTED_SCALE_STOP");
        if (rule.column === undefined || rule.when !== undefined || rule.target === "row") unsupported("UNSUPPORTED_SCALE_TARGET");
        numeric(rule.column!);
        const { stops } = captureFields(styleInput.scale, ["stops"]);
        if (!Array.isArray(stops) || stops.length < 2 || stops.length > 3) unsupported("UNSUPPORTED_SCALE_STOPS");
        const mapped = stops.map((input: NonNullable<CanopyFormatRule["style"]["scale"]>["stops"][number]) => {
          const stop = captureFields(input, ["type", "value", "tone"]);
          if (!["min", "max", "number", "percent", "percentile"].includes(stop.type)) unsupported("UNSUPPORTED_SCALE_THRESHOLD:" + stop.type);
          const threshold: ConditionalThreshold = stop.type === "min" || stop.type === "max" ? { type: stop.type }
            : { type: stop.type as "number" | "percent" | "percentile", value: stop.value! };
          return { threshold, color: color(stop.tone) };
        });
        visuals.push({ type: "colorScale", range: { column: rule.column! },
          stops: mapped as [ConditionalColorStop, ConditionalColorStop] | [ConditionalColorStop, ConditionalColorStop, ConditionalColorStop] });
        requiredCells.set(column(rule.column!), "number");
      } else {
        if (rule.when === undefined || styleInput.tone === undefined || rule.target !== "row" && rule.column === undefined) unsupported("UNSUPPORTED_HIGHLIGHT_TARGET");
        if (rule.target !== undefined && rule.target !== "row" && rule.target !== "cell") unsupported("UNSUPPORTED_RULE_TARGET");
        const range = rule.target === "row" ? { column: columns[0]!.id, through: columns.at(-1)!.id } : { column: rule.column! };
        if (rule.column !== undefined) column(rule.column);
        const ruleCells = new Map<number, "number" | "date">();
        const ownPredicate = canopyPredicate(request, grouped, options, rule.when!, unsupported, ruleCells);
        const reachable = priorStops.length ? "AND(" + ownPredicate + ",NOT(OR(" + priorStops.join(",") + ")))" : ownPredicate;
        const previous = rule.target === "row" ? priorRows : priorCells.get(rule.column!) ?? [];
        const formula = previous.length ? "AND(" + reachable + ",NOT(OR(" + previous.join(",") + ")))" : reachable;
        if (formula.length > 8192) unsupported("RULE_FORMULA_LIMIT");
        (rule.target === "row" ? rows : cells).push({ type: "expression", range, formula, style: style(styleInput.tone!) });
        previous.push(ownPredicate);
        if (rule.target !== "row") priorCells.set(rule.column!, previous);
        for (const [column, type] of ruleCells) requiredCells.set(column, type);
        // A matched declaration still stops the row when an earlier tone on that target already won.
        if (rule.stopIfTrue) priorStops.push(ownPredicate);
      }
    } catch (error) {
      if (!(error instanceof UnsupportedRule)) throw error;
      report({ code: error.message + ":" + rule.id, ...(rule.column === undefined ? {} : { columnId: rule.column }) });
      if (hasStops) return []; // A partial sequence cannot claim to preserve its row-global stopping behavior.
    }
  }
  for (const [column, type] of requiredCells) cellTypes?.set(column, type);
  // Canopy paints cell visuals/tones over the row tint. Build masks before changing native priority order.
  return Object.freeze([...visuals, ...cells, ...rows]);
}
