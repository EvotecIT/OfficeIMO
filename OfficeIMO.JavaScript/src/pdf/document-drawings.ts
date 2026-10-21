import { NotSupportedError, OfficeIMOError } from "../core/errors.js";
import { checkAbort, pause, taskYieldDue } from "../core/iteration.js";
import type { PdfDrawing, PdfDrawingElement, PdfDrawingPath, PdfDrawingText, PdfDrawingClip, PdfDrawingFont } from "./drawing-types.js";
import type { PdfFontResource, PdfFontResources } from "./font-resources.js";
import { color, positive } from "./settings.js";
import type { PdfSettings } from "./settings.js";
import { pdfNumber as n } from "./objects.js";
import type { PdfPathCommand } from "./document-types.js";

type PathPlan = { readonly kind: "path"; readonly commands: readonly PdfPathCommand[]; readonly style: string; readonly paint: string };
type TextPlan = { readonly kind: "text"; readonly text: string; readonly x: number; readonly y: number; readonly size: number; readonly color: string; readonly font: PdfDrawingFont };
type GroupPlan = { readonly kind: "group"; readonly clip?: PdfDrawingClip; readonly children: readonly Plan[] };
type Plan = PathPlan | TextPlan | GroupPlan;
export interface PdfDrawingUsage { readonly commands: number; readonly elements: number; }
export interface PreparedPdfDrawing extends PdfDrawingUsage { readonly width: number; readonly height: number; readonly content: string; }
interface CapturedDrawing extends PdfDrawingUsage { readonly width: number; readonly height: number; readonly plan: readonly Plan[]; }

/** Portable drawings are an exact schema: unsupported data must never be silently discarded. */
function record(value: unknown, keys: readonly string[], name: string): Record<string, unknown> {
  if (!value || typeof value !== "object" || Array.isArray(value) || Reflect.ownKeys(value).some(key => typeof key !== "string" || !keys.includes(key)))
    throw new TypeError("Unsupported " + name + " properties.");
  return value as Record<string, unknown>;
}
export function checkDrawingBlock(value: unknown): void {
  record(value, ["kind", "drawing", "alignment", "alternativeText", "decorative", "caption"], "drawing block");
}
function coordinate(value: unknown): number {
  if (typeof value !== "number" || !Number.isFinite(value) || Math.abs(value) > 1000000)
    throw new RangeError("Drawing coordinates must be finite and within -1,000,000 through 1,000,000 points.");
  return value;
}
function clipRectangle(value: unknown): PdfDrawingClip {
  const clip = record(value, ["x", "y", "width", "height"], "drawing clip");
  return { x: coordinate(clip.x), y: coordinate(clip.y), width: positive(clip.width as number, "Drawing clip width", 1000000), height: positive(clip.height as number, "Drawing clip height", 1000000) };
}
async function cooperate(settings: PdfSettings): Promise<void> {
  if (taskYieldDue()) await pause();
  checkAbort(settings.options.signal);
}

/** Take ownership of one bounded schema before any cooperative or font work can run caller code. */
function captureDrawing(drawing: PdfDrawing, settings: PdfSettings, used: PdfDrawingUsage): CapturedDrawing {
  checkAbort(settings.options.signal);
  const root = record(drawing, ["version", "width", "height", "elements"], "drawing");
  if (root.version !== 1) throw new NotSupportedError("Only portable drawing version 1 is supported.");
  const width = positive(root.width as number, "Drawing width"), height = positive(root.height as number, "Drawing height");
  let commands = used.commands, elements = used.elements;
  const ancestors = new Set<object>();
  function path(value: PdfDrawingPath): PathPlan {
    record(value, ["kind", "commands", "fill", "stroke", "strokeWidth", "fillRule", "lineCap", "lineJoin"], "drawing path");
    const source = value.commands;
    if (!Array.isArray(source)) throw new TypeError("Drawing path commands must be an array.");
    if (commands + source.length > settings.limits.maxVectorCommands) throw new OfficeIMOError("RESOURCE_LIMIT", "maxVectorCommands exceeded.");
    checkAbort(settings.options.signal);
    const snapshot = source.slice(), copied: PdfPathCommand[] = []; let started = false;
    commands += snapshot.length;
    const fillValue = value.fill, strokeValue = value.stroke;
    const fill = fillValue === undefined ? undefined : color(fillValue), stroke = strokeValue === undefined ? undefined : color(strokeValue);
    if (fill === undefined && stroke === undefined) throw new TypeError("A drawing path requires fill or stroke.");
    const strokeWidth = positive(value.strokeWidth ?? 1, "Drawing stroke width");
    const fillRule = value.fillRule, lineCap = value.lineCap, lineJoin = value.lineJoin;
    if (fillRule !== undefined && !["nonzero", "evenodd"].includes(fillRule)) throw new TypeError("Invalid drawing fillRule.");
    if (lineCap !== undefined && !["butt", "round", "square"].includes(lineCap)) throw new TypeError("Invalid drawing lineCap.");
    if (lineJoin !== undefined && !["miter", "round", "bevel"].includes(lineJoin)) throw new TypeError("Invalid drawing lineJoin.");
    for (let i = 0; i < snapshot.length; i++) {
      const c = snapshot[i]!;
      if (!c || typeof c !== "object") throw new TypeError("Invalid drawing path command.");
      const op = c.op;
      if (op === "move" || op === "line") {
        record(c, ["op", "x", "y"], "drawing path command");
        if (op === "line" && !started) throw new TypeError("A drawing subpath must begin with move.");
        copied.push({ op, x: coordinate(c.x), y: coordinate(c.y) }); started = true;
      } else if (op === "cubic") {
        record(c, ["op", "x1", "y1", "x2", "y2", "x", "y"], "drawing path command");
        if (!started) throw new TypeError("A drawing subpath must begin with move.");
        copied.push({ op: "cubic", x1: coordinate(c.x1), y1: coordinate(c.y1), x2: coordinate(c.x2), y2: coordinate(c.y2), x: coordinate(c.x), y: coordinate(c.y) });
      } else if (op === "close") {
        record(c, ["op"], "drawing path command");
        if (!started) throw new TypeError("Cannot close an empty drawing path."); copied.push({ op: "close" });
      } else throw new TypeError("Unsupported drawing path command.");
      if ((i & 1023) === 1023) checkAbort(settings.options.signal);
    }
    if (!started) throw new TypeError("A drawing path needs at least one command.");
    return { kind: "path", commands: copied, style: (fill === undefined ? "" : fill + " rg\n") + (stroke === undefined ? "" : stroke + " RG\n" + n(strokeWidth) + " w\n") +
      String(["butt", "round", "square"].indexOf(lineCap ?? "butt")) + " J\n" + String(["miter", "round", "bevel"].indexOf(lineJoin ?? "miter")) + " j\n10 M\n",
      paint: (fill !== undefined && stroke !== undefined ? "B" : fill !== undefined ? "f" : "S") + (fill !== undefined && fillRule === "evenodd" ? "*" : "") + "\n" };
  }
  function text(value: PdfDrawingText): TextPlan {
    record(value, ["kind", "text", "x", "y", "font", "fontSize", "color"], "drawing text");
    const snapshot = value.text, slot = value.font;
    if (typeof snapshot !== "string") throw new TypeError("Drawing text must be a string.");
    if (snapshot.length > settings.limits.maxCellCharacters) throw new OfficeIMOError("RESOURCE_LIMIT", "maxCellCharacters exceeded.");
    if (/[\r\n\u2028\u2029]/.test(snapshot)) throw new NotSupportedError("Drawing text must be a single baseline without line breaks.");
    if (!["regular", "bold", "italic", "boldItalic"].includes(slot)) throw new TypeError("Invalid drawing font slot.");
    const x = coordinate(value.x), y = coordinate(value.y), size = positive(value.fontSize, "Drawing font size"), foreground = color(value.color);
    settings.budget.cell(snapshot);
    return { kind: "text", text: snapshot, x, y, size, color: foreground, font: slot };
  }
  function children(input: readonly PdfDrawingElement[], depth: number): readonly Plan[] {
    if (!Array.isArray(input)) throw new TypeError("Drawing elements must be an array.");
    if (depth > 32 && input.length) throw new RangeError("Drawing groups support at most 32 element levels.");
    if (elements + input.length > settings.limits.maxDrawingElements) throw new OfficeIMOError("RESOURCE_LIMIT", "maxDrawingElements exceeded.");
    checkAbort(settings.options.signal);
    const snapshot = input.slice(), output: Plan[] = [];
    for (let i = 0; i < snapshot.length; i++) {
      if (++elements > settings.limits.maxDrawingElements) throw new OfficeIMOError("RESOURCE_LIMIT", "maxDrawingElements exceeded.");
      const value = snapshot[i];
      if (!value || typeof value !== "object") throw new TypeError("Invalid drawing element.");
      if (ancestors.has(value)) throw new TypeError("Drawing groups cannot contain cycles.");
      const kind = value.kind;
      if (kind === "path") output.push(path(value));
      else if (kind === "text") output.push(text(value));
      else if (kind === "group") {
        record(value, ["kind", "clip", "elements"], "drawing group");
        const clipValue = value.clip, clip = clipValue === undefined ? undefined : clipRectangle(clipValue);
        ancestors.add(value);
        output.push({ kind: "group", ...(clip ? { clip } : {}), children: children(value.elements, depth + 1) });
        ancestors.delete(value);
      } else throw new TypeError("Unsupported drawing element kind.");
      if ((i & 255) === 255) checkAbort(settings.options.signal);
    }
    return output;
  }
  const plan = children(root.elements as readonly PdfDrawingElement[], 1);
  checkAbort(settings.options.signal);
  return { width, height, plan, commands: commands - used.commands, elements: elements - used.elements };
}

/** Preflight and paint owned records. The existing font owner retains glyph maps and subsets. */
export async function prepareDrawing(drawing: PdfDrawing, settings: PdfSettings, fonts: PdfFontResources,
  used: PdfDrawingUsage): Promise<PreparedPdfDrawing> {
  // A drawing's bounded capture is synchronous; yielding here would leave later caller records borrowed.
  const captured = captureDrawing(drawing, settings, used), drawingFonts = new Map<PdfDrawingFont, PdfFontResource>();
  function fontFor(slot: PdfDrawingFont): PdfFontResource {
    const existing = drawingFonts.get(slot);
    if (existing) return existing;
    const font = fonts.select(slot === "bold" || slot === "boldItalic", slot === "italic" || slot === "boldItalic");
    if (!font.program || font.syntheticBold || font.syntheticItalic) throw new NotSupportedError("Drawing text without an exact embedded font face at fonts." + slot);
    drawingFonts.set(slot, font); return font;
  }
  // The work budget belongs to the drawing, including short runs in separate sibling groups.
  let preflightCharacters = 0, preflightElements = 0;
  async function preflight(nodes: readonly Plan[]): Promise<void> {
    for (let i = 0; i < nodes.length; i++) {
      const node = nodes[i]!;
      if (node.kind === "group") await preflight(node.children);
      else if (node.kind === "text") {
        const font = fontFor(node.font);
        for (const character of node.text) {
          font.width(character.codePointAt(0)!);
          if ((++preflightCharacters & 1023) === 0) await cooperate(settings);
        }
      }
      if ((++preflightElements & 255) === 0) await cooperate(settings);
    }
  }
  await preflight(captured.plan);
  const fragments: string[] = []; let bytes = 0;
  function add(fragment: string): void {
    bytes += fragment.length;
    if (bytes > settings.limits.maxPageBytes - 256) throw new OfficeIMOError("RESOURCE_LIMIT", "maxPageBytes exceeded by drawing content.");
    fragments.push(fragment);
  }
  let paintedCharacters = 0, paintedCommands = 0, paintedElements = 0;
  async function paint(nodes: readonly Plan[]): Promise<void> {
    for (let i = 0; i < nodes.length; i++) {
      const node = nodes[i]!; add("q\n");
      if (node.kind === "group") {
        if (node.clip) add([node.clip.x, node.clip.y, node.clip.width, node.clip.height].map(n).join(" ") + " re W n\n");
        await paint(node.children);
      } else if (node.kind === "path") {
        add(node.style);
        for (let j = 0; j < node.commands.length; j++) {
          const c = node.commands[j]!;
          add(c.op === "move" || c.op === "line" ? n(c.x) + " " + n(c.y) + (c.op === "move" ? " m\n" : " l\n") :
            c.op === "cubic" ? [c.x1, c.y1, c.x2, c.y2, c.x, c.y].map(n).join(" ") + " c\n" : "h\n");
          if ((++paintedCommands & 1023) === 0) await cooperate(settings);
        }
        add(node.paint);
      } else {
        const font = fontFor(node.font);
        add(node.color + " rg\nBT /" + font.name + " " + n(node.size) + " Tf 0 Tr 1 0 0 -1 " + n(node.x) + " " + n(node.y) + " Tm <");
        let chunk = "";
        for (const character of node.text) {
          chunk += character;
          if ((++paintedCharacters & 1023) === 0) { add(font.encode(chunk).slice(1, -1)); chunk = ""; await cooperate(settings); }
        }
        if (chunk) add(font.encode(chunk).slice(1, -1));
        add("> Tj ET\n");
      }
      add("Q\n");
      if ((++paintedElements & 255) === 0) await cooperate(settings);
    }
  }
  await paint(captured.plan); checkAbort(settings.options.signal);
  return { width: captured.width, height: captured.height, content: fragments.join(""), commands: captured.commands, elements: captured.elements };
}
