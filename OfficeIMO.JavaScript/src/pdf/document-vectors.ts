import { OfficeIMOError } from "../core/errors.js";
import { checkAbort, pause, taskYieldDue } from "../core/iteration.js";
import type { PdfVectorBlock } from "./document-types.js";
import { color, positive } from "./settings.js";
import { pdfNumber as n } from "./objects.js";
import type { PdfSettings } from "./settings.js";

interface PreparedVector { readonly width: number; readonly height: number; readonly count: number; readonly content: string; }
/** Capture and serialize one bounded display path, with no markup parser or resource loading. */
export async function vectorCommands(block: PdfVectorBlock, settings: PdfSettings, used: number): Promise<PreparedVector> {
  const width = positive(block.width, "Vector width"), height = positive(block.height, "Vector height");
  const source = block.commands;
  if (!Array.isArray(source)) throw new TypeError("Vector commands must be an array.");
  if (used + source.length > settings.limits.maxVectorCommands) throw new OfficeIMOError("RESOURCE_LIMIT", "maxVectorCommands exceeded.");
  const input = source.map(c => {
    if (!c || typeof c !== "object") throw new TypeError("Invalid vector command.");
    const op = c.op;
    if (op === "move" || op === "line") return { op, x: c.x, y: c.y };
    if (op === "cubic") return { op, x1: c.x1, y1: c.y1, x2: c.x2, y2: c.y2, x: c.x, y: c.y };
    if (op === "close") return { op };
    throw new TypeError("Unsupported vector command.");
  });
  const fill = block.fill === undefined ? undefined : color(block.fill), stroke = block.stroke === undefined ? undefined : color(block.stroke);
  if (fill === undefined && stroke === undefined) throw new TypeError("A vector path requires fill or stroke.");
  const strokeWidth = positive(block.strokeWidth ?? 1, "Vector stroke width", 144);
  const fillRule = block.fillRule;
  if (fillRule !== undefined && !["nonzero", "evenodd"].includes(fillRule)) throw new TypeError("Invalid vector fill rule.");
  const commands: string[] = [];
  let bytes = 0, started = false;
  const coordinate = (value: number, bound: number): string => {
    if (!Number.isFinite(value) || value < 0 || value > bound) throw new RangeError("Vector coordinates must stay inside their declared bounds.");
    return n(value);
  };
  for (let i = 0; i < input.length; i++) {
    const c = input[i]!; let command: string;
    if (!c || typeof c !== "object") throw new TypeError("Invalid vector command.");
    if (c.op === "move" || c.op === "line") {
      if (c.op === "line" && !started) throw new TypeError("A vector subpath must begin with move.");
      command = coordinate(c.x, width) + " " + coordinate(c.y, height) + (c.op === "move" ? " m\n" : " l\n"); started = true;
    } else if (c.op === "cubic") {
      if (!started) throw new TypeError("A vector subpath must begin with move.");
      command = [coordinate(c.x1, width), coordinate(c.y1, height), coordinate(c.x2, width), coordinate(c.y2, height), coordinate(c.x, width), coordinate(c.y, height)].join(" ") + " c\n";
    } else if (c.op === "close") { if (!started) throw new TypeError("Cannot close an empty vector path."); command = "h\n"; }
    else throw new TypeError("Unsupported vector command.");
    bytes += command.length;
    if (bytes > settings.limits.maxPageBytes - 256) throw new OfficeIMOError("RESOURCE_LIMIT", "maxPageBytes exceeded by vector commands.");
    commands.push(command);
    if ((i & 1023) === 1023) { if (taskYieldDue()) await pause(); checkAbort(settings.options.signal); }
  }
  if (!started) throw new TypeError("A vector path needs at least one command.");
  return { width, height, count: input.length, content: (fill === undefined ? "" : fill + " rg\n") + (stroke === undefined ? "" : stroke + " RG\n" + n(strokeWidth) + " w\n") +
    commands.join("") + (fill !== undefined && stroke !== undefined ? "B" : fill !== undefined ? "f" : "S") + (fill !== undefined && fillRule === "evenodd" ? "*" : "") + "\n" };
}
