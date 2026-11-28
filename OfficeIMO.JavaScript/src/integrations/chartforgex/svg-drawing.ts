import { NotSupportedError, OfficeIMOError } from "../../core/errors.js";
import { checkAbort } from "../../core/iteration.js";
import { pdfNumber } from "../../pdf/objects.js";
import type { PdfDrawing, PdfDrawingElement, PdfDrawingFont, PdfDrawingPath, PdfDrawingText } from "../../pdf/drawing-types.js";
import { capturePdfLimits, pdfDrawingLimits } from "../../pdf/settings.js";
import { parseSvgMarkup } from "./svg-markup.js";
import type { SvgNode } from "./svg-markup.js";
import { apply, ellipse, identity, multiply, numbers, parsePathData, parseTransform, polyline, rectangle, scaleOf, transformCommands } from "./svg-geometry.js";
import type { Matrix, UserCommand } from "./svg-geometry.js";
import { alpha, flatten, length, parseColor, parseStyle } from "./svg-paint.js";
import type { Rgba } from "./svg-paint.js";
import type { ChartForgeXDiagnostic, ChartForgeXSvgOptions } from "./types.js";
import { winAnsiText } from "./winansi.js";

const inherited = new Set(["fill", "fill-opacity", "fill-rule", "stroke", "stroke-opacity", "stroke-width", "stroke-linecap", "stroke-linejoin",
  "stroke-dasharray", "stroke-dashoffset", "font-size", "font-weight", "font-style", "text-anchor", "dominant-baseline", "alignment-baseline",
  "color", "visibility", "white-space"]);
const own = ["opacity", "display", "clip-path", "mask", "filter", "marker-start", "marker-mid", "marker-end", "vector-effect", "transform"] as const;
const skipped = new Set(["defs", "clipPath", "mask", "marker", "pattern", "linearGradient", "radialGradient", "symbol", "script", "title", "desc",
  "metadata", "filter", "stop", "font", "font-face"]);
const shapes = new Set(["path", "rect", "circle", "ellipse", "line", "polyline", "polygon"]);
const initial: ReadonlyMap<string, string> = new Map([["fill", "black"], ["stroke", "none"], ["stroke-width", "1"], ["font-size", "16"], ["color", "black"]]);

interface Context {
  readonly matrix: Matrix;
  readonly userWidth: number;
  readonly userHeight: number;
  readonly style: ReadonlyMap<string, string>;
  readonly fontSize: number;
  readonly opacity: number;
}
type ViewportSize = { readonly width: number | undefined; readonly height: number | undefined };
const nonUniform = (m: Matrix): boolean => {
  const tolerance = 1e-6, xScale = Math.hypot(m[0], m[1]), yScale = Math.hypot(m[2], m[3]);
  return Math.abs(xScale - yScale) > tolerance * Math.max(xScale, yScale) ||
    Math.abs(m[0] * m[2] + m[1] * m[3]) > tolerance * xScale * yScale;
};

/** Convert a resolved chart SVG into the PDF writer's portable drawing. Lengths become points (96 px = 72 pt). */
export function svgToPdfDrawing(svg: string, options: ChartForgeXSvgOptions = {}, signal?: AbortSignal): PdfDrawing {
  const limits = options.limits === undefined ? undefined : capturePdfLimits(options.limits);
  const maxElements = limits?.maxDrawingElements ?? pdfDrawingLimits.maxDrawingElements;
  const maxCommands = limits?.maxVectorCommands ?? pdfDrawingLimits.maxVectorCommands;
  const maxWork = Math.max(4096, 16 * Math.max(maxElements, maxCommands));
  let elementCount = 0, commandCount = 0, work = 0;
  const charge = (amount = 1): void => {
    checkAbort(signal);
    work += amount;
    if (work > maxWork) throw new OfficeIMOError("RESOURCE_LIMIT", "Chart SVG conversion work limit exceeded.");
  };
  const addElement = (target: PdfDrawingElement[], element: PdfDrawingElement): void => {
    if (++elementCount > maxElements) throw new OfficeIMOError("RESOURCE_LIMIT", "maxDrawingElements exceeded.");
    target.push(element);
  };
  if (typeof svg === "string") charge(Math.ceil(svg.length / 16));
  const root = parseSvgMarkup(svg), ids = new Map<string, SvgNode>(), reported = new Set<string>();
  const fonts = new Set<PdfDrawingFont>(options.fonts ?? ["regular", "bold", "italic", "boldItalic"]), winAnsi = options.winAnsi !== false;
  const index = (node: SvgNode): void => {
    charge();
    const id = node.attributes.get("id");
    if (id !== undefined && !ids.has(id)) ids.set(id, node);
    for (const child of node.children) if (typeof child !== "string") index(child);
  };
  index(root);
  // Skippable effects are reported once per code; reject turns the first into an error.
  const diagnose = (code: string, element?: string): void => {
    if (options.unsupported === "reject") throw new NotSupportedError("Unsupported chart SVG feature: " + code + (element ? " (<" + element + ">)" : ""));
    if (reported.has(code)) return;
    reported.add(code);
    const diagnostic: ChartForgeXDiagnostic = element === undefined ? { code } : { code, element };
    options.onDiagnostic?.(diagnostic);
  };
  const attribute = (node: SvgNode, name: string): string | undefined => node.attributes.get(name);

  const viewport = (node: SvgNode, fontSize: number, fallbackWidth: number, fallbackHeight: number, nested = false, size?: ViewportSize) => {
    const box = numbers(attribute(node, "viewBox") ?? ""), hasBox = box.length === 4 && box[2]! > 0 && box[3]! > 0;
    const referenceWidth = nested || !hasBox ? fallbackWidth : box[2]!, referenceHeight = nested || !hasBox ? fallbackHeight : box[3]!;
    const width = size?.width ?? length(attribute(node, "width"), fontSize, referenceWidth) ?? referenceWidth;
    const height = size?.height ?? length(attribute(node, "height"), fontSize, referenceHeight) ?? referenceHeight;
    if (!hasBox) return { width, height, userWidth: width, userHeight: height, matrix: identity };
    const aspect = (attribute(node, "preserveAspectRatio") ?? "xMidYMid meet").trim().split(/\s+/);
    let sx = width / box[2]!, sy = height / box[3]!, tx = 0, ty = 0;
    if (aspect[0] !== "none") {
      const scale = aspect[1] === "slice" ? Math.max(sx, sy) : Math.min(sx, sy), align = aspect[0] ?? "xMidYMid";
      tx = (width - box[2]! * scale) * (align.includes("xMid") ? .5 : align.includes("xMax") ? 1 : 0);
      ty = (height - box[3]! * scale) * (align.includes("YMid") ? .5 : align.includes("YMax") ? 1 : 0);
      sx = sy = scale;
    }
    const matrix: Matrix = [sx, 0, 0, sy, tx - box[0]! * sx, ty - box[1]! * sy];
    return { width, height, userWidth: box[2]!, userHeight: box[3]!, matrix };
  };

  const outer = viewport(root, 16, 0, 0);
  if (!(outer.width > 0 && outer.height > 0)) throw new TypeError("Chart SVG needs a positive width and height or viewBox.");
  // Points at 96 CSS px per inch, reduced further only when the drawing would exceed the PDF drawing bounds.
  const unit = Math.min(.75, 14400 / outer.width, 14400 / outer.height);
  const elements: PdfDrawingElement[] = [];

  function cascade(node: SvgNode, parent: Context): { context: Context; own: Map<string, string> } | undefined {
    const inline = attribute(node, "style");
    if (inline) charge(Math.ceil(inline.length / 16));
    const declared = parseStyle(inline), values = new Map<string, string>();
    let style = parent.style, changed = false;
    const read = (name: string): string | undefined => declared.get(name) ?? attribute(node, name);
    for (const name of inherited) {
      const value = read(name);
      if (value === undefined || value === "inherit" || value === parent.style.get(name)) continue;
      if (!changed) { style = new Map(parent.style); changed = true; }
      (style as Map<string, string>).set(name, value);
    }
    for (const name of own) { const value = read(name); if (value !== undefined) values.set(name, value); }
    if (values.get("display") === "none") return undefined;
    const size = read("font-size");
    const fontSize = size !== undefined && size !== "inherit" ? fontSizeOf(size, parent.fontSize) : parent.fontSize;
    const transform = parseTransform(values.get("transform"));
    if (!transform) diagnose("TRANSFORM_UNSUPPORTED", node.local);
    const opacity = alpha(values.get("opacity")) ?? 1;
    if (opacity < 1 && opacity > 0) diagnose("OPACITY_FLATTENED", node.local);
    return { context: { ...parent, matrix: multiply(parent.matrix, transform ?? identity), style, fontSize, opacity: parent.opacity * opacity }, own: values };
  }
  function fontSizeOf(value: string, parent: number): number {
    const keywords: Readonly<Record<string, number>> = { "xx-small": 9, "x-small": 10, small: 13, medium: 16, large: 18, "x-large": 24, "xx-large": 32 };
    const size = keywords[value.trim()] ?? length(value, parent, parent);
    return size !== undefined && size > 0 ? size : parent;
  }

  function paint(value: string | undefined, context: Context, opacityName: string, element: string): string | undefined {
    if (value === undefined || value === "none") return undefined;
    let color: Rgba | undefined;
    const reference = /^url\(\s*["']?#([^"')]+)["']?\s*\)\s*(.*)$/.exec(value.trim());
    if (reference) {
      const target = ids.get(reference[1]!), fallback = reference[2]!.trim();
      if (target && (target.local === "linearGradient" || target.local === "radialGradient")) {
        color = gradientColor(target);
        diagnose("GRADIENT_FLATTENED", element);
      } else if (fallback && fallback !== "none") color = parseColor(fallback);
      else { diagnose("PAINT_UNSUPPORTED", element); return undefined; }
    } else if (value.trim().toLowerCase() === "currentcolor") color = parseColor(context.style.get("color") ?? "black");
    else color = parseColor(value);
    if (!color) { diagnose("COLOR_UNSUPPORTED", element); return undefined; }
    const effective = (alpha(context.style.get(opacityName)) ?? 1) * context.opacity;
    if (color.a * effective < 1 && color.a * effective >= .002) diagnose("OPACITY_FLATTENED", element);
    return flatten(color, effective);
  }
  function gradientColor(node: SvgNode): Rgba | undefined {
    const seen = new Set<SvgNode>();
    let current: SvgNode | undefined = node;
    while (current && !seen.has(current)) {
      charge();
      seen.add(current);
      for (const child of current.children) {
        charge();
        if (typeof child === "string" || child.local !== "stop") continue;
        const declared = parseStyle(attribute(child, "style")), color = parseColor(declared.get("stop-color") ?? attribute(child, "stop-color") ?? "black");
        const opacity = alpha(declared.get("stop-opacity") ?? attribute(child, "stop-opacity")) ?? 1;
        return color && { ...color, a: color.a * opacity };
      }
      const href: string | undefined = (attribute(current, "href") ?? attribute(current, "xlink:href"))?.replace(/^#/, "");
      current = href === undefined ? undefined : ids.get(href);
    }
    return undefined;
  }

  // Consecutive paths with identical paint merge into one element, which keeps dense scatter and bar charts within drawing
  // budgets. Only strokes and simple shapes of one winding merge, so a nonzero fill never gains a hole.
  const mergeKeys = new WeakMap<PdfDrawingElement, string>();
  function pushPath(target: PdfDrawingElement[], path: PdfDrawingPath, key: string | undefined): void {
    const last = target[target.length - 1];
    if (key !== undefined && last && last.kind === "path" && mergeKeys.get(last) === key && last.commands.length + path.commands.length <= 4096) {
      (last.commands as PdfDrawingPath["commands"][number][]).push(...path.commands); return;
    }
    if (key !== undefined) mergeKeys.set(path, key);
    addElement(target, path);
  }

  function shape(node: SvgNode, context: Context, target: PdfDrawingElement[], values: ReadonlyMap<string, string>): void {
    const { userWidth, userHeight } = context;
    const diagonal = Math.hypot(userWidth, userHeight) / Math.SQRT2;
    const a = (name: string, reference = 0): number => length(attribute(node, name), context.fontSize, reference) ?? 0;
    let commands: UserCommand[];
    switch (node.local) {
      case "rect": {
        const width = a("width", userWidth), height = a("height", userHeight);
        if (!(width > 0 && height > 0)) return;
        const rxValue = attribute(node, "rx"), ryValue = attribute(node, "ry");
        const rx = rxValue === undefined || rxValue === "auto" ? a("ry", userHeight) : a("rx", userWidth);
        const ry = ryValue === undefined || ryValue === "auto" ? rx : a("ry", userHeight);
        commands = rectangle(a("x", userWidth), a("y", userHeight), width, height, rx, ry); break;
      }
      case "circle": { const r = a("r", diagonal); if (!(r > 0)) return; commands = ellipse(a("cx", userWidth), a("cy", userHeight), r, r); break; }
      case "ellipse": { const rx = a("rx", userWidth), ry = a("ry", userHeight); if (!(rx > 0 && ry > 0)) return; commands = ellipse(a("cx", userWidth), a("cy", userHeight), rx, ry); break; }
      case "line": commands = [{ op: "move", x: a("x1", userWidth), y: a("y1", userHeight) }, { op: "line", x: a("x2", userWidth), y: a("y2", userHeight) }]; break;
      case "polyline": case "polygon": {
        const points = attribute(node, "points") ?? "";
        charge(Math.ceil(points.length / 4));
        commands = polyline(numbers(points), node.local === "polygon"); break;
      }
      default: {
        const data = attribute(node, "d") ?? "";
        charge(Math.ceil(data.length / 4));
        const parsed = parsePathData(data);
        if (parsed.malformed) diagnose("PATH_DATA_TRUNCATED", "path");
        commands = parsed.commands;
      }
    }
    if (!commands.length) return;
    for (const marker of ["marker-start", "marker-mid", "marker-end"]) if (values.get(marker) && values.get(marker) !== "none") diagnose("MARKER_IGNORED", node.local);
    const fill = node.local === "line" ? undefined : paint(context.style.get("fill"), context, "fill-opacity", node.local);
    const stroke = paint(context.style.get("stroke"), context, "stroke-opacity", node.local);
    if (fill === undefined && stroke === undefined) return;
    const fixedStroke = values.get("vector-effect") === "non-scaling-stroke";
    const scale = fixedStroke ? unit : scaleOf(context.matrix);
    const width = (length(context.style.get("stroke-width"), context.fontSize, diagonal) ?? 1) * scale;
    if (stroke !== undefined && !(width > 0) && fill === undefined) return;
    commandCount += commands.length;
    if (commandCount > maxCommands) throw new OfficeIMOError("RESOURCE_LIMIT", "maxVectorCommands exceeded.");
    const path: { -readonly [K in keyof PdfDrawingPath]: PdfDrawingPath[K] } = { kind: "path", commands: transformCommands(commands, context.matrix) };
    if (fill !== undefined) { path.fill = fill; if (context.style.get("fill-rule") === "evenodd") path.fillRule = "evenodd"; }
    if (stroke !== undefined && width > 0) {
      if (!fixedStroke && nonUniform(context.matrix)) diagnose("STROKE_DISTORTED", node.local);
      path.stroke = stroke; path.strokeWidth = width;
      const cap = context.style.get("stroke-linecap"), join = context.style.get("stroke-linejoin");
      if (cap === "round" || cap === "square") path.lineCap = cap;
      if (join === "round" || join === "bevel") path.lineJoin = join;
      const dash = (context.style.get("stroke-dasharray") ?? "none").trim();
      if (dash !== "none" && dash) {
        const parts = dash.split(/[\s,]+/).map(part => length(part, context.fontSize, diagonal));
        if (parts.every(part => part !== undefined && part >= 0) && parts.some(part => part! > 0)) {
          const list = (parts.length % 2 ? [...parts, ...parts] : parts).map(part => part! * scale);
          const offset = (length(context.style.get("stroke-dashoffset"), context.fontSize, diagonal) ?? 0) * scale;
          if (list.length <= 16 && list.every(part => part <= 14400) && list.some(part => pdfNumber(part) !== "0") && Math.abs(offset) <= 1000000) {
            path.dash = list;
            if (offset) path.dashOffset = offset;
          } else diagnose("DASH_UNSUPPORTED", node.local);
        } else if (parts.some(part => part === undefined || part < 0)) diagnose("DASH_UNSUPPORTED", node.local);
      }
    }
    if (path.fill === undefined && path.stroke === undefined) return;
    const mergeable = path.fill === undefined ||
      (path.stroke === undefined || path.stroke === path.fill) && path.fillRule === undefined && ["rect", "circle", "ellipse"].includes(node.local);
    const det = context.matrix[0] * context.matrix[3] - context.matrix[1] * context.matrix[2];
    pushPath(target, path, mergeable ? JSON.stringify([path.fill, det < 0, path.stroke, path.strokeWidth, path.lineCap, path.lineJoin, path.dash, path.dashOffset]) : undefined);
  }

  function textRuns(node: SvgNode, context: Context, target: PdfDrawingElement[]): void {
    // A run starts at the text element and at each explicitly positioned tspan; unpositioned tspans continue the run in its
    // starting style, because the drawing cannot measure where earlier text ends.
    type Run = { x: number; y: number; text: string; context: Context };
    const runs: Run[] = [];
    let x = 0, y = 0, current: Run | undefined;
    const walk = (element: SvgNode, at: Context, first: boolean): void => {
      charge();
      const list = (name: string, reference: number): number[] => (attribute(element, name) ?? "").trim().split(/[\s,]+/).filter(Boolean)
        .map(item => length(item, at.fontSize, reference)).filter((item): item is number => item !== undefined);
      const xs = list("x", at.userWidth), ys = list("y", at.userHeight), dxs = list("dx", at.userWidth), dys = list("dy", at.userHeight);
      if (xs.length > 1 || ys.length > 1 || dxs.length > 1 || dys.length > 1 || attribute(element, "rotate") !== undefined) diagnose("TEXT_POSITIONS_SIMPLIFIED", element.local);
      const positioned = first || xs.length > 0 || ys.length > 0 || dxs.length > 0 || dys.length > 0;
      // Without x, a move is relative to where the unmeasured earlier text ends.
      if (!first && positioned && !xs.length && current?.text.trim()) diagnose("TEXT_POSITIONS_SIMPLIFIED", element.local);
      if (positioned) {
        if (xs.length) x = xs[0]!;
        if (ys.length) y = ys[0]!;
        x += dxs[0] ?? 0; y += dys[0] ?? 0;
        current = { x, y, text: "", context: at }; runs.push(current);
      }
      for (const child of element.children) {
        if (typeof child === "string") {
          charge(1 + Math.ceil(child.length / 16));
          if (at.style.get("visibility") === "hidden" || at.style.get("visibility") === "collapse" || at.opacity <= 0 || !current) continue;
          const from = current.context;
          if ((from.style !== at.style || from.fontSize !== at.fontSize || from.opacity !== at.opacity) && /\S/.test(child))
            diagnose("TEXT_STYLE_MERGED", element.local);
          current.text += child;
          continue;
        }
        if (child.local !== "tspan" && child.local !== "a") { if (child.local === "textPath") diagnose("TEXT_PATH_IGNORED", "textPath"); continue; }
        const nested = cascade(child, at);
        if (nested) walk(child, nested.context, false);
      }
    };
    walk(node, context, true);
    for (const run of runs) {
      const style = run.context.style, preserve = attribute(node, "xml:space") === "preserve" || /^pre/.test(style.get("white-space") ?? "");
      let text = run.text.replace(/[\r\n\t]/g, preserve ? " " : "");
      if (!preserve) text = text.replace(/ {2,}/g, " ");
      if (run === runs[0] && !preserve) text = text.replace(/^ +/, "");
      if (run === runs[runs.length - 1] && !preserve) text = text.replace(/ +$/, "");
      if (!text) continue;
      if (style.get("visibility") === "hidden" || style.get("visibility") === "collapse") continue;
      const color = paint(style.get("fill"), run.context, "fill-opacity", "text");
      if (color === undefined) continue;
      if ((style.get("stroke") ?? "none") !== "none") diagnose("TEXT_STROKE_IGNORED", "text");
      const m = run.context.matrix, scale = scaleOf(m);
      if (nonUniform(m)) diagnose("TEXT_DISTORTED", "text");
      const size = run.context.fontSize * scale;
      if (!(size > 0)) continue;
      const baseline = style.get("dominant-baseline") ?? style.get("alignment-baseline") ?? "auto";
      const shift = baseline === "middle" || baseline === "central" ? .35 : baseline === "hanging" || baseline === "text-before-edge" || baseline === "text-top" ? .8 :
        baseline === "text-after-edge" || baseline === "ideographic" || baseline === "text-bottom" ? -.2 : 0;
      const [px, py] = apply(m, run.x, run.y + shift * run.context.fontSize), rotation = Math.atan2(m[1], m[0]) * 180 / Math.PI;
      const weight = style.get("font-weight") ?? "normal", italic = /italic|oblique/.test(style.get("font-style") ?? "");
      const bold = weight === "bold" || weight === "bolder" || Number(weight) >= 600;
      let slot: PdfDrawingFont = bold && italic ? "boldItalic" : bold ? "bold" : italic ? "italic" : "regular";
      if (!fonts.has(slot)) slot = "regular";
      const anchor = style.get("text-anchor");
      const element: { -readonly [K in keyof PdfDrawingText]: PdfDrawingText[K] } = { kind: "text", text: winAnsi ? winAnsiText(text, () => diagnose("TEXT_SUBSTITUTED", "text")) : text, x: px, y: py, font: slot,
        fontSize: size, color };
      if (!element.text) continue;
      if (anchor === "middle" || anchor === "end") element.anchor = anchor;
      if (Math.abs(rotation) > 1e-6) element.rotation = Math.round(rotation * 1e6) / 1e6;
      addElement(target, element);
    }
  }
  function hasText(node: SvgNode): boolean {
    return node.children.some(child => {
      charge(typeof child === "string" ? 1 + Math.ceil(child.length / 16) : 1);
      return typeof child === "string" ? /\S/.test(child) : hasText(child);
    });
  }
  function clipFor(value: string | undefined, context: Context, element: string): { x: number; y: number; width: number; height: number } | null | undefined {
    if (value === undefined || value === "none") return undefined;
    const id = /^url\(\s*["']?#([^"')]+)["']?\s*\)$/.exec(value.trim())?.[1], clip = id === undefined ? undefined : ids.get(id);
    let rect: SvgNode | undefined;
    if (clip?.local === "clipPath") for (const child of clip.children) {
      charge();
      if (typeof child === "string") continue;
      if (rect || child.local !== "rect") { rect = undefined; break; }
      rect = child;
    }
    const m = multiply(context.matrix, parseTransform(clip?.attributes.get("transform")) ?? identity);
    if (!rect || clip!.attributes.get("clipPathUnits") === "objectBoundingBox" || Math.abs(m[1]) > 1e-9 || Math.abs(m[2]) > 1e-9) {
      diagnose("CLIP_IGNORED", element); return null;
    }
    const r = multiply(m, parseTransform(rect.attributes.get("transform")) ?? identity);
    if (Math.abs(r[1]) > 1e-9 || Math.abs(r[2]) > 1e-9) { diagnose("CLIP_IGNORED", element); return null; }
    const a = (name: string, reference: number): number => length(rect.attributes.get(name), context.fontSize, reference) ?? 0;
    const x = a("x", context.userWidth), y = a("y", context.userHeight);
    const [x1, y1] = apply(r, x, y), [x2, y2] = apply(r, x + a("width", context.userWidth), y + a("height", context.userHeight));
    return { x: Math.min(x1, x2), y: Math.min(y1, y2), width: Math.max(Math.abs(x2 - x1), 1e-3), height: Math.max(Math.abs(y2 - y1), 1e-3) };
  }

  const active = new Set<SvgNode>();
  function render(node: SvgNode, parent: Context, target: PdfDrawingElement[], depth: number, useSize?: ViewportSize): void {
    charge();
    if (skipped.has(node.local) && !(node.local === "symbol" && useSize)) return;
    if (depth > 512) throw new NotSupportedError("Chart SVG nesting is too deep.");
    if (node.local === "image") { diagnose("IMAGE_SKIPPED", "image"); return; }
    if (node.local === "foreignObject") { diagnose("FOREIGN_OBJECT_SKIPPED", "foreignObject"); return; }
    if (node.local === "style") { if (hasText(node)) diagnose("STYLESHEET_IGNORED", "style"); return; }
    const cascaded = cascade(node, parent);
    if (!cascaded || cascaded.context.opacity <= 0) return;
    let { context } = cascaded;
    const values = cascaded.own;
    if (values.get("mask") && values.get("mask") !== "none") diagnose("MASK_IGNORED", node.local);
    if (values.get("filter") && values.get("filter") !== "none") diagnose("FILTER_IGNORED", node.local);
    if ((node.local === "svg" || node.local === "symbol") && node !== root) {
      const nested = viewport(node, context.fontSize, context.userWidth, context.userHeight, true, useSize);
      if (!(nested.width > 0 && nested.height > 0)) return;
      const x = length(attribute(node, "x"), context.fontSize, context.userWidth) ?? 0, y = length(attribute(node, "y"), context.fontSize, context.userHeight) ?? 0;
      context = { ...context, userWidth: nested.userWidth, userHeight: nested.userHeight, matrix: multiply(context.matrix, multiply([1, 0, 0, 1, x, y], nested.matrix)) };
    }
    const clip = clipFor(values.get("clip-path"), context, node.local);
    let output = target;
    if (clip) { const children: PdfDrawingElement[] = []; addElement(target, { kind: "group", clip, elements: children }); output = children; }
    const visible = context.style.get("visibility") !== "hidden" && context.style.get("visibility") !== "collapse";
    if (shapes.has(node.local)) { if (visible) shape(node, context, output, values); }
    else if (node.local === "text") textRuns(node, context, output);
    else if (node.local === "use") {
      const href = (attribute(node, "href") ?? attribute(node, "xlink:href"))?.replace(/^#/, ""), referenced = href === undefined ? undefined : ids.get(href);
      if (!referenced || active.has(referenced)) diagnose("USE_UNRESOLVED", "use");
      else {
        const x = length(attribute(node, "x"), context.fontSize, context.userWidth) ?? 0, y = length(attribute(node, "y"), context.fontSize, context.userHeight) ?? 0;
        const placed = { ...context, matrix: multiply(context.matrix, [1, 0, 0, 1, x, y]) };
        active.add(referenced);
        const useSize = { width: length(attribute(node, "width"), context.fontSize, context.userWidth),
          height: length(attribute(node, "height"), context.fontSize, context.userHeight) };
        render(referenced, placed, output, depth + 1, useSize);
        active.delete(referenced);
      }
    } else if (node.local === "switch") {
      for (const child of node.children) {
        charge();
        if (typeof child === "string") continue;
        if (["systemLanguage", "requiredExtensions", "requiredFeatures"].some(name => child.attributes.has(name))) {
          diagnose("SWITCH_CONDITION_IGNORED", child.local);
          continue;
        }
        render(child, context, output, depth + 1);
        break;
      }
    } else if (["svg", "symbol", "g", "a"].includes(node.local)) {
      for (const child of node.children) { if (typeof child === "string") charge(); else render(child, context, output, depth + 1); }
    } else diagnose("ELEMENT_SKIPPED", node.local);
    if (clip && !output.length) { target.pop(); elementCount--; }
  }
  render(root, { matrix: [unit * outer.matrix[0], 0, 0, unit * outer.matrix[3], unit * outer.matrix[4], unit * outer.matrix[5]],
    userWidth: outer.userWidth, userHeight: outer.userHeight, style: initial, fontSize: 16, opacity: 1 }, elements, 0);
  return { version: 1, width: Math.min(14400, outer.width * unit), height: Math.min(14400, outer.height * unit), elements, standardFonts: true };
}
