import { OfficeIMOError } from "../core/errors.js";
import { cleanXml, escapeOoxmlAttribute } from "../xml/index.js";
import type { InvalidCharacterPolicy } from "../xml/index.js";
import { colorArgb } from "./style-xml.js";
import type { ConditionalStyle } from "./conditional-types.js";

/** @internal Reject unsupported properties rather than silently losing presentation. */
export function fields(value: unknown, allowed: readonly string[], label: string): void {
  if (!value || typeof value !== "object" || Array.isArray(value)) throw new TypeError(label + " must be an object.");
  if (typeof (value as { then?: unknown }).then === "function") {
    void Promise.resolve(value).catch(() => {}); throw new TypeError(label + " must be synchronous.");
  }
  for (const key of Object.keys(value)) if (!allowed.includes(key)) throw new TypeError("Unsupported " + label + " property: " + key);
}
/** @internal */
export function optionalBoolean(value: unknown, name: string): void {
  if (value !== undefined && typeof value !== "boolean") throw new TypeError(name + " must be boolean.");
}
/** @internal Differential styles contain no default font/fill/border properties. */
export class DifferentialStyles {
  private readonly definitions = new Map<string, { id: number; xml: string }>();
  constructor(private readonly policy: InvalidCharacterPolicy, private readonly maximum: number, private readonly numberFormat: (format: string) => number) {}
  private keys(styles: readonly ConditionalStyle[]): readonly string[] {
    // Validation and canonical XML precede registration, so equivalent property order shares a dxf.
    const keys = styles.map(style => differentialXml(style, this.policy, () => 164));
    const pending = new Set(keys.filter(definition => !this.definitions.has(definition)));
    if (this.definitions.size + pending.size > this.maximum) throw new OfficeIMOError("RESOURCE_LIMIT", "maxDifferentialStyles exceeded.");
    return keys;
  }
  check(styles: readonly ConditionalStyle[]): void { this.keys(styles); }
  add(styles: readonly ConditionalStyle[]): readonly number[] {
    const keys = this.keys(styles);
    for (let i = 0; i < styles.length; i++) if (!this.definitions.has(keys[i]!))
      this.definitions.set(keys[i]!, { id: this.definitions.size, xml: differentialXml(styles[i]!, this.policy, this.numberFormat) });
    return keys.map(key => this.definitions.get(key)!.id);
  }
  toXml(): string {
    return this.definitions.size ? '<dxfs count="' + this.definitions.size + '">' + [...this.definitions.values()].map(style => '<dxf>' + style.xml + '</dxf>').join("") + '</dxfs>' : "";
  }
}
/** @internal Also used for preflight before registering a worksheet name. */
export function differentialXml(style: ConditionalStyle, policy: InvalidCharacterPolicy, numberFormat: (format: string) => number): string {
  fields(style, ["font", "fill", "border", "numberFormat"], "conditional style");
  let xml = "";
  if (style.font !== undefined) {
    fields(style.font, ["bold", "italic", "underline", "strike", "color"], "conditional font");
    let font = "";
    for (const [name, tag] of [["bold", "b"], ["italic", "i"], ["strike", "strike"]] as const) {
      optionalBoolean(style.font[name], name);
      if (style.font[name] !== undefined) font += '<' + tag + ' val="' + (style.font[name] ? 1 : 0) + '"/>';
    }
    optionalBoolean(style.font.underline, "underline");
    if (style.font.underline !== undefined) font += '<u val="' + (style.font.underline ? "single" : "none") + '"/>';
    if (style.font.color !== undefined) font += '<color rgb="' + colorArgb(style.font.color) + '"/>';
    if (font) xml += '<font>' + font + '</font>';
  }
  if (style.numberFormat !== undefined) {
    const format = style.numberFormat;
    if (typeof format !== "string" || !format || format.length > 255 || !cleanXml(format, policy)) throw new TypeError("Conditional number format requires 1 through 255 XML-valid characters.");
    xml += '<numFmt numFmtId="' + numberFormat(format) + '" formatCode="' + escapeOoxmlAttribute(format, policy) + '"/>';
  }
  if (style.fill !== undefined) {
    fields(style.fill, ["color"], "conditional fill");
    const color = colorArgb(style.fill.color);
    // Excel reads the background color for a differential solid fill; keep both channels explicit.
    xml += '<fill><patternFill patternType="solid"><fgColor rgb="' + color + '"/><bgColor rgb="' + color + '"/></patternFill></fill>';
  }
  if (style.border !== undefined) {
    fields(style.border, ["left", "right", "top", "bottom"], "conditional border");
    let border = "";
    for (const side of ["left", "right", "top", "bottom"] as const) {
      const edge = style.border[side];
      if (edge !== undefined) {
        fields(edge, ["style", "color"], "conditional border edge");
        if (!["thin", "medium", "dashed", "dotted", "thick", "double", "hair", "dashDot", "dashDotDot", "slantDashDot", "mediumDashed", "mediumDashDot", "mediumDashDotDot"].includes(edge.style)) throw new TypeError("Invalid conditional border style.");
        border += '<' + side + ' style="' + edge.style + '">' + (edge.color !== undefined ? '<color rgb="' + colorArgb(edge.color) + '"/>' : "") + '</' + side + '>';
      }
    }
    if (border) xml += '<border>' + border + '</border>';
  }
  if (!xml) throw new TypeError("A conditional style must change at least one property.");
  return xml;
}
