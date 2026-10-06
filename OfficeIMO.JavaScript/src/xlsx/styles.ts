import { escapeOoxmlAttribute, cleanXml, xmlDeclaration } from "../xml/index.js";
import type { InvalidCharacterPolicy } from "../xml/index.js";
import type { Alignment, Column } from "../core/index.js";
import { OfficeIMOError } from "../core/errors.js";
export const spreadsheetNamespace = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";

export interface Font { readonly name?: string; readonly size?: number; readonly bold?: boolean; readonly italic?: boolean; readonly underline?: boolean; readonly strike?: boolean; readonly color?: string; }
export interface Fill { readonly pattern?: "none" | "gray125" | "solid"; readonly color?: string; }
export type BorderLineStyle = "thin" | "medium" | "dashed" | "dotted" | "thick" | "double" | "hair" | "dashDot" | "dashDotDot" | "slantDashDot" | "mediumDashed" | "mediumDashDot" | "mediumDashDotDot";
export interface BorderEdge { readonly style: BorderLineStyle; readonly color?: string; }
export interface Border { readonly left?: BorderEdge; readonly right?: BorderEdge; readonly top?: BorderEdge; readonly bottom?: BorderEdge; }
export interface CellStyle {
  readonly font?: number | Font;
  readonly fill?: number | Fill;
  readonly border?: number | Border;
  readonly numberFormat?: string | number;
  readonly wrapText?: boolean;
  readonly alignment?: Alignment;
  readonly verticalAlignment?: "top" | "center" | "bottom" | "justify" | "distributed";
}
export const NumberFormats = { General: "General", Integer: "0", Decimal: "0.00", Percent: "0.00%", Date: "yyyy-mm-dd", DateTime: "yyyy-mm-dd hh:mm:ss" } as const;

export function colorArgb(value: string): string {
  if (typeof value !== "string" || !/^#?(?:[0-9a-f]{6}|[0-9a-f]{8})$/i.test(value)) throw new TypeError("Color must be RGB or ARGB hex.");
  const color = value.replace(/^#/, "").toUpperCase(); return color.length === 6 ? "FF" + color : color;
}
function index(value: number, count: number, kind: string): number {
  if (!Number.isInteger(value) || value < 0 || value >= count) throw new RangeError("Unknown " + kind + " index.");
  return value;
}
function deduplicate<T>(items: T[], indexes: Map<string, number>, item: T, maximum: number): number {
  const key = JSON.stringify(item), found = indexes.get(key);
  if (found !== undefined) return found;
  if (items.length >= maximum) throw new OfficeIMOError("RESOURCE_LIMIT", "Style component limit exceeded.");
  const id = items.length; items.push(item); indexes.set(key, id); return id;
}
interface RegisteredStyle { font: number; fill: number; border: number; numberFormat: number; wrapText: boolean; alignment: string; verticalAlignment: string; }

/** @internal Observe unsupported promises immediately, even when cell precedence skips the patch. */
export function validateStylePatch(patch: CellStyle): void {
  if (!patch || typeof patch !== "object" || Array.isArray(patch)) throw new TypeError("A style patch must be a CellStyle object.");
  if (typeof (patch as { then?: unknown }).then === "function") {
    void Promise.resolve(patch).catch(() => {});
    throw new TypeError("Style patches must be synchronous CellStyle objects.");
  }
}
/** @internal Capture reusable option patches without retaining caller-owned nested components. */
export function copyStylePatch(patch: CellStyle): CellStyle {
  validateStylePatch(patch);
  return Object.freeze({ ...patch,
    ...(typeof patch.font === "object" ? { font: Object.freeze({ ...patch.font }) } : {}),
    ...(typeof patch.fill === "object" ? { fill: Object.freeze({ ...patch.fill }) } : {}),
    ...(typeof patch.border === "object" ? { border: Object.freeze(Object.fromEntries(Object.entries(patch.border).map(([side, edge]) => [side, Object.freeze({ ...edge })]))) } : {})
  });
}

/** Workbook-owned indexes. Definition inputs are normalized and copied, never retained by reference. */
export class StyleRegistry {
  private readonly fonts: Font[] = [{ name: "Calibri", size: 11, bold: false, italic: false, underline: false, strike: false, color: "" }];
  private readonly fills: Fill[] = [{ pattern: "none", color: "" }, { pattern: "gray125", color: "" }];
  private readonly borders: Border[] = [{}];
  private readonly formats = new Map<string, number>();
  private readonly styles: RegisteredStyle[] = [];
  private readonly indexes = new Map<string, number>();
  private readonly fontIndexes = new Map(this.fonts.map((v, i) => [JSON.stringify(v), i]));
  private readonly fillIndexes = new Map(this.fills.map((v, i) => [JSON.stringify(v), i]));
  private readonly borderIndexes = new Map(this.borders.map((v, i) => [JSON.stringify(v), i]));
  constructor(private readonly policy: InvalidCharacterPolicy = "strip", private readonly maximum = 64000) {
    if (!Number.isSafeInteger(maximum) || maximum < 1 || maximum > 64000) throw new RangeError("maxStyles must be from 1 through 64,000.");
    cleanXml("", policy); this.add();
  }
  addFont(font: Font): number {
    if (font.name !== undefined && typeof font.name !== "string") throw new TypeError("Font name must be a string.");
    const name = cleanXml(font.name ?? "Calibri", this.policy), size = font.size ?? 11;
    if (typeof name !== "string" || !name || name.length > 31) throw new TypeError("Font name must contain 1 through 31 characters.");
    if (!Number.isFinite(size) || size <= 0 || size > 409) throw new RangeError("Font size must be positive and at most 409.");
    return deduplicate(this.fonts, this.fontIndexes, { name, size, bold: !!font.bold, italic: !!font.italic, underline: !!font.underline, strike: !!font.strike, color: font.color ? colorArgb(font.color) : "" }, this.maximum);
  }
  addFill(fill: Fill): number {
    const pattern = fill.pattern ?? (fill.color ? "solid" : "none");
    if (!["none", "gray125", "solid"].includes(pattern) || (pattern === "solid" && !fill.color) || (pattern !== "solid" && fill.color))
      throw new TypeError("A solid fill requires a color; none and gray125 have no color.");
    return deduplicate(this.fills, this.fillIndexes, { pattern, color: fill.color ? colorArgb(fill.color) : "" }, this.maximum + 1);
  }
  addBorder(border: Border): number {
    const result: { left?: BorderEdge; right?: BorderEdge; top?: BorderEdge; bottom?: BorderEdge } = {};
    for (const side of ["left", "right", "top", "bottom"] as const) {
      const edge = border[side];
      if (edge) {
        if (!["thin", "medium", "dashed", "dotted", "thick", "double", "hair", "dashDot", "dashDotDot", "slantDashDot", "mediumDashed", "mediumDashDot", "mediumDashDotDot"].includes(edge.style))
          throw new TypeError("Invalid border style.");
        result[side] = { style: edge.style, ...(edge.color ? { color: colorArgb(edge.color) } : {}) };
      }
    }
    return deduplicate(this.borders, this.borderIndexes, result, this.maximum);
  }
  addNumberFormat(format: string): number {
    if (typeof format !== "string" || !format || format.length > 255) throw new TypeError("Number format must contain 1 through 255 characters.");
    format = cleanXml(format, this.policy);
    if (!format) throw new TypeError("Number format must contain XML-valid characters.");
    if (format === "General") return 0;
    let id = this.formats.get(format);
    if (id === undefined) {
      if (this.formats.size >= 65372) throw new RangeError("Too many number formats.");
      id = this.formats.size + 164; this.formats.set(format, id);
    }
    return id;
  }
  add(style: CellStyle = {}): number {
    const font = typeof style.font === "number" ? index(style.font, this.fonts.length, "font") : style.font ? this.addFont(style.font) : 0;
    const fill = typeof style.fill === "number" ? index(style.fill, this.fills.length, "fill") : style.fill ? this.addFill(style.fill) : 0;
    const border = typeof style.border === "number" ? index(style.border, this.borders.length, "border") : style.border ? this.addBorder(style.border) : 0;
    let numberFormat = 0;
    if (typeof style.numberFormat === "string") numberFormat = this.addNumberFormat(style.numberFormat);
    else if (style.numberFormat !== undefined) {
      numberFormat = style.numberFormat;
      if (!Number.isInteger(numberFormat) || numberFormat < 0 || (numberFormat > 163 && ![...this.formats.values()].includes(numberFormat)))
        throw new RangeError("Unknown number format index.");
    }
    if (style.alignment !== undefined && !["left", "center", "right", "fill", "justify", "distributed"].includes(style.alignment)) throw new TypeError("Invalid horizontal alignment.");
    if (style.verticalAlignment !== undefined && !["top", "center", "bottom", "justify", "distributed"].includes(style.verticalAlignment)) throw new TypeError("Invalid vertical alignment.");
    const registered = { font, fill, border, numberFormat, wrapText: !!style.wrapText, alignment: style.alignment ?? "", verticalAlignment: style.verticalAlignment ?? "" };
    const key = JSON.stringify(registered), found = this.indexes.get(key);
    if (found !== undefined) return found;
    if (this.styles.length >= this.maximum) throw new OfficeIMOError("RESOURCE_LIMIT", "Workbook exceeds its cell style limit.");
    const id = this.styles.length; this.styles.push(registered); this.indexes.set(key, id); return id;
  }
  validateStyle(id: number): number { return index(id, this.styles.length, "cell style"); }
  /** Overlay presentation while retaining the base number format, font fields and border edges. Numeric component indexes replace a component. */
  compose(base: number, overlay: CellStyle): number {
    validateStylePatch(overlay);
    const source = this.styles[this.validateStyle(base)]!;
    return this.add({
      font: source.font, fill: source.fill, border: source.border, numberFormat: source.numberFormat,
      wrapText: source.wrapText,
      ...(source.alignment ? { alignment: source.alignment as Alignment } : {}),
      ...(source.verticalAlignment ? { verticalAlignment: source.verticalAlignment as CellStyle["verticalAlignment"] & string } : {}),
      ...overlay,
      ...(typeof overlay.font === "object" ? { font: { ...this.fonts[source.font], ...overlay.font } } : {}),
      ...(typeof overlay.border === "object" ? { border: { ...this.borders[source.border], ...overlay.border } } : {})
    });
  }
  /** @internal Column shorthand and header style composition. */
  forColumn(column: Column, header = false, fill?: string, date = false): number {
    const base = !header && column.style !== undefined ? this.styles[this.validateStyle(column.style)] : undefined;
    const format = column.format !== undefined ? column.format || NumberFormats.General : base?.numberFormat || (column.type === "date" || date ? NumberFormats.DateTime : NumberFormats.General);
    return this.add({
      ...(base ? { font: base.font, fill: base.fill, border: base.border } : {}),
      numberFormat: format,
      wrapText: column.wrapText ?? base?.wrapText ?? false,
      ...(column.alignment ? { alignment: column.alignment } : base?.alignment ? { alignment: base.alignment as Alignment } : {}),
      ...(base?.verticalAlignment ? { verticalAlignment: base.verticalAlignment as CellStyle["verticalAlignment"] & string } : {}),
      ...(header ? { font: { bold: true } } : {}), ...(fill ? { fill: { color: fill } } : {})
    });
  }
  toXml(): string {
    const fontXml = (f: Font) => '<font>' + (f.bold ? '<b/>' : "") + (f.italic ? '<i/>' : "") + (f.underline ? '<u/>' : "") + (f.strike ? '<strike/>' : "") + '<sz val="' + f.size + '"/>' +
      (f.color ? '<color rgb="' + f.color + '"/>' : "") + '<name val="' + escapeOoxmlAttribute(f.name, this.policy) + '"/></font>';
    const borderXml = (b: Border) => '<border>' + (["left", "right", "top", "bottom"] as const).map(side => {
      const edge = b[side]; return edge ? '<' + side + ' style="' + edge.style + '">' + (edge.color ? '<color rgb="' + edge.color + '"/>' : "") + '</' + side + '>' : '<' + side + '/>';
    }).join("") + '<diagonal/></border>';
    return xmlDeclaration + '<styleSheet xmlns="' + spreadsheetNamespace + '">' +
      '<numFmts count="' + this.formats.size + '">' + [...this.formats].map(([format, id]) => '<numFmt numFmtId="' + id + '" formatCode="' + escapeOoxmlAttribute(format, this.policy) + '"/>').join("") + '</numFmts>' +
      '<fonts count="' + this.fonts.length + '">' + this.fonts.map(fontXml).join("") + '</fonts>' +
      '<fills count="' + this.fills.length + '">' + this.fills.map(f => '<fill><patternFill patternType="' + f.pattern + '">' +
        (f.color ? '<fgColor rgb="' + f.color + '"/><bgColor indexed="64"/>' : "") + '</patternFill></fill>').join("") + '</fills>' +
      '<borders count="' + this.borders.length + '">' + this.borders.map(borderXml).join("") + '</borders>' +
      '<cellStyleXfs count="1"><xf numFmtId="0" fontId="0" fillId="0" borderId="0"/></cellStyleXfs>' +
      '<cellXfs count="' + this.styles.length + '">' + this.styles.map(s => {
        const alignment = s.wrapText || s.alignment || s.verticalAlignment ? '<alignment' + (s.wrapText ? ' wrapText="1"' : "") + (s.alignment ? ' horizontal="' + s.alignment + '"' : "") + (s.verticalAlignment ? ' vertical="' + s.verticalAlignment + '"' : "") + '/>' : "";
        return '<xf numFmtId="' + s.numberFormat + '" fontId="' + s.font + '" fillId="' + s.fill + '" borderId="' + s.border + '" xfId="0"' +
          (s.numberFormat ? ' applyNumberFormat="1"' : "") + (s.font ? ' applyFont="1"' : "") + (s.fill ? ' applyFill="1"' : "") +
          (s.border ? ' applyBorder="1"' : "") + (alignment ? ' applyAlignment="1"' : "") + '>' + alignment + '</xf>';
      }).join("") + '</cellXfs><cellStyles count="1"><cellStyle name="Normal" xfId="0" builtinId="0"/></cellStyles></styleSheet>';
  }
}
