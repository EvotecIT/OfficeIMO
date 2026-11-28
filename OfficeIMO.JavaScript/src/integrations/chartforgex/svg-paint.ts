/** Straight RGBA, channels 0-255 and alpha 0-1. */
export interface Rgba { readonly r: number; readonly g: number; readonly b: number; readonly a: number; }

const named: Readonly<Record<string, string>> = {
  black: "000000", white: "ffffff", red: "ff0000", green: "008000", blue: "0000ff", yellow: "ffff00", orange: "ffa500", purple: "800080",
  gray: "808080", grey: "808080", silver: "c0c0c0", maroon: "800000", navy: "000080", teal: "008080", olive: "808000", lime: "00ff00",
  aqua: "00ffff", cyan: "00ffff", fuchsia: "ff00ff", magenta: "ff00ff", darkgray: "a9a9a9", darkgrey: "a9a9a9", lightgray: "d3d3d3",
  lightgrey: "d3d3d3", dimgray: "696969", dimgrey: "696969", gainsboro: "dcdcdc", whitesmoke: "f5f5f5", steelblue: "4682b4",
  slategray: "708090", slategrey: "708090", crimson: "dc143c", tomato: "ff6347", gold: "ffd700", indigo: "4b0082", brown: "a52a2a",
  pink: "ffc0cb", coral: "ff7f50", salmon: "fa8072", seagreen: "2e8b57", forestgreen: "228b22", royalblue: "4169e1", darkblue: "00008b",
  darkred: "8b0000", darkgreen: "006400", darkorange: "ff8c00", lightblue: "add8e6", skyblue: "87ceeb", tan: "d2b48c", violet: "ee82ee"
};

const clamp = (value: number, low: number, high: number): number => Math.min(high, Math.max(low, value));
function channel(text: string, scale: number): number {
  const value = parseFloat(text);
  return text.trim().endsWith("%") ? value / 100 * scale : value;
}
export function alpha(text: string | undefined): number | undefined {
  if (text === undefined) return 1;
  if (text.trim() === "none") return 0;
  if (!/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?%?$/i.test(text.trim())) return undefined;
  return clamp(channel(text, 1), 0, 1);
}
function functionArguments(body: string, legacyAlpha: boolean): { readonly values: string[]; readonly alpha: string | undefined } {
  const [main, slash] = body.split("/") as [string, string | undefined];
  const values = main.includes(",") ? main.split(",").map(s => s.trim()).filter(Boolean) : main.trim().split(/\s+/);
  // Legacy rgba(r, g, b, a) and hsla carry alpha as a fourth argument; color(space c1 c2 c3) does not.
  if (legacyAlpha && slash === undefined && values.length === 4) return { values: values.slice(0, 3), alpha: values[3] };
  return { values, alpha: slash?.trim() };
}
function linearToSrgb(value: number): number {
  const v = clamp(value, 0, 1);
  return 255 * (v <= 0.0031308 ? 12.92 * v : 1.055 * Math.pow(v, 1 / 2.4) - 0.055);
}
function fromOklab(l: number, a: number, b: number, opacity: number): Rgba {
  const l1 = Math.pow(l + 0.3963377774 * a + 0.2158037573 * b, 3), m1 = Math.pow(l - 0.1055613458 * a - 0.0638541728 * b, 3);
  const s1 = Math.pow(l - 0.0894841775 * a - 1.2914855480 * b, 3);
  return { r: linearToSrgb(4.0767416621 * l1 - 3.3077115913 * m1 + 0.2309699292 * s1),
    g: linearToSrgb(-1.2684380046 * l1 + 2.6097574011 * m1 - 0.3413193965 * s1),
    b: linearToSrgb(-0.0041960863 * l1 - 0.7034186147 * m1 + 1.7076147010 * s1), a: opacity };
}
function hue(text: string): number {
  const value = parseFloat(text);
  return text.endsWith("turn") ? value * 360 : text.endsWith("rad") ? value * 180 / Math.PI : value;
}

/** Parse the CSS color forms a browser's computed style and SVG presentation attributes produce. */
export function parseColor(input: string): Rgba | undefined {
  const text = input.trim().toLowerCase();
  if (!text) return undefined;
  if (text === "transparent") return { r: 0, g: 0, b: 0, a: 0 };
  const hex = /^#([0-9a-f]{3,8})$/.exec(text)?.[1];
  if (hex) {
    if (hex.length === 3 || hex.length === 4) {
      const [r, g, b, a] = [...hex].map(c => parseInt(c + c, 16));
      return { r: r!, g: g!, b: b!, a: a === undefined ? 1 : a / 255 };
    }
    if (hex.length === 6 || hex.length === 8)
      return { r: parseInt(hex.slice(0, 2), 16), g: parseInt(hex.slice(2, 4), 16), b: parseInt(hex.slice(4, 6), 16), a: hex.length === 8 ? parseInt(hex.slice(6), 16) / 255 : 1 };
    return undefined;
  }
  const known = named[text];
  if (known) return parseColor("#" + known);
  const call = /^([a-z]+)\((.*)\)$/.exec(text);
  if (!call) return undefined;
  const name = call[1]!, { values, alpha: opacity } = functionArguments(call[2]!, name !== "color");
  if (values.some(value => !/^[+-]?(?:\d+\.?\d*|\.\d+)(?:e[+-]?\d+)?(%|deg|rad|turn)?$/.test(value) && value !== "none")) {
    if (name !== "color") return undefined;
  }
  const a = alpha(opacity);
  if (a === undefined) return undefined;
  const v = values.map(value => value === "none" ? "0" : value);
  if ((name === "rgb" || name === "rgba") && v.length === 3)
    return { r: clamp(channel(v[0]!, 255), 0, 255), g: clamp(channel(v[1]!, 255), 0, 255), b: clamp(channel(v[2]!, 255), 0, 255), a };
  if ((name === "hsl" || name === "hsla") && v.length === 3) {
    const h = ((hue(v[0]!) % 360) + 360) % 360 / 360, s = clamp(parseFloat(v[1]!) / 100, 0, 1), l = clamp(parseFloat(v[2]!) / 100, 0, 1);
    const q = l < .5 ? l * (1 + s) : l + s - l * s, p = 2 * l - q;
    const tone = (t: number): number => { t = (t + 1) % 1; return 255 * (t < 1 / 6 ? p + (q - p) * 6 * t : t < .5 ? q : t < 2 / 3 ? p + (q - p) * (2 / 3 - t) * 6 : p); };
    return { r: tone(h + 1 / 3), g: tone(h), b: tone(h - 1 / 3), a };
  }
  if (name === "oklab" && v.length === 3) return fromOklab(channel(v[0]!, 1), channel(v[1]!, .4), channel(v[2]!, .4), a);
  if (name === "oklch" && v.length === 3) {
    const l = channel(v[0]!, 1), c = channel(v[1]!, .4), h = hue(v[2]!) * Math.PI / 180;
    return fromOklab(l, c * Math.cos(h), c * Math.sin(h), a);
  }
  if (name === "color" && v.length === 4 && (v[0] === "srgb" || v[0] === "srgb-linear")) {
    const parts = v.slice(1).map(value => channel(value, 1));
    const convert = v[0] === "srgb" ? (x: number): number => clamp(x, 0, 1) * 255 : linearToSrgb;
    if (!parts.every(Number.isFinite)) return undefined;
    return { r: convert(parts[0]!), g: convert(parts[1]!), b: convert(parts[2]!), a };
  }
  return undefined;
}

/** Composite over white, the page colour, and return six-digit hex for the PDF drawing contract. */
export function flatten(color: Rgba, opacity: number): string | undefined {
  const a = clamp(color.a * opacity, 0, 1);
  if (a < 0.002) return undefined;
  const mix = (value: number): string => Math.round(clamp(value, 0, 255) * a + 255 * (1 - a)).toString(16).padStart(2, "0").toUpperCase();
  return mix(color.r) + mix(color.g) + mix(color.b);
}

/** Inline `style` declarations as a map; later declarations win, `!important` is accepted and dropped. */
export function parseStyle(text: string | undefined): Map<string, string> {
  const result = new Map<string, string>();
  if (!text) return result;
  const important = new Set<string>();
  for (const declaration of text.split(";")) {
    const colon = declaration.indexOf(":");
    if (colon < 0) continue;
    const name = declaration.slice(0, colon).trim().toLowerCase(), raw = declaration.slice(colon + 1);
    const priority = /!\s*important\s*$/i;
    const value = raw.replace(priority, "").trim(), takesPriority = priority.test(raw);
    if (!name || !value || important.has(name) && !takesPriority) continue;
    if (takesPriority) important.add(name);
    result.set(name, value);
  }
  return result;
}

/** A CSS/SVG length in user units. Percentages resolve against `reference`; `em` against the font size. */
export function length(text: string | undefined, fontSize: number, reference = 0): number | undefined {
  if (text === undefined) return undefined;
  const match = /^\s*([+-]?(?:\d+\.?\d*|\.\d+)(?:[eE][+-]?\d+)?)\s*(px|pt|pc|mm|cm|in|em|rem|ex|%)?\s*$/.exec(text);
  if (!match) return undefined;
  const value = Number(match[1]), unit = match[2];
  const factor = unit === undefined || unit === "px" ? 1 : unit === "pt" ? 4 / 3 : unit === "pc" ? 16 : unit === "mm" ? 96 / 25.4 : unit === "cm" ? 96 / 2.54 :
    unit === "in" ? 96 : unit === "em" ? fontSize : unit === "rem" ? 16 : unit === "ex" ? fontSize / 2 : reference / 100;
  const result = value * factor;
  return Number.isFinite(result) ? result : undefined;
}
