import { NotSupportedError, OfficeIMOError } from "../core/errors.js";
import { copyExportLink } from "../core/links.js";
import { checkAbort, pause, taskYieldDue } from "../core/iteration.js";
import type { CellPresentation, ExportLink } from "../core/index.js";
import type { PdfFontResource, PdfFontResources } from "./font-resources.js";
import type { PdfText, PdfTextRun } from "./inline-types.js";
import type { PdfSettings } from "./settings.js";
import { presentation, color } from "./settings.js";
import { wrapShapedPdfText } from "./shaped-text.js";
import type { PdfShapedGlyph } from "./text-shaping-types.js";
import { PdfFont } from "./font.js";
import { pdfBudgetText } from "./text.js";

export interface CapturedRun {
  readonly text: string;
  readonly style: Readonly<CellPresentation>;
  readonly link?: Readonly<ExportLink>;
  readonly font?: PdfFont;
}
/** @internal Immutable input captured before layout can yield to application code. */
export interface PdfCapturedText { readonly text: string; readonly runs: readonly CapturedRun[]; }
export interface PdfInlineSegment {
  readonly text: string;
  readonly width: number;
  readonly font: PdfFontResource;
  readonly color?: string;
  readonly link?: Readonly<ExportLink>;
  readonly glyphs?: readonly (PdfShapedGlyph & { readonly logicalText: string })[];
  readonly logicalStart?: number;
  readonly actualText?: string;
}
export interface PdfInlineLine {
  readonly text: string;
  readonly width: number;
  readonly segments: readonly PdfInlineSegment[];
  readonly lineHeight: number;
  readonly ascent: number;
  readonly actualText?: string;
}

/** @internal Callers supply an owned, validated, frozen base presentation. A run list is a PDF presentation value, never an alternate shared table value model. */
export function capturePdfText(value: PdfText, base: Readonly<CellPresentation>, settings: PdfSettings, fonts: PdfFontResources, countEmpty = true, reserveBudget = true): PdfCapturedText {
  if (typeof value === "string") {
    if (value.length > settings.limits.maxCellCharacters) throw new OfficeIMOError("RESOURCE_LIMIT", "maxCellCharacters exceeded.");
    if (reserveBudget && (value.length || countEmpty)) settings.budget.reserve(1, pdfBudgetText(value, settings.budget).length);
    return Object.freeze({ text: value, runs: Object.freeze([Object.freeze({ text: value, style: base })]) });
  }
  const input: readonly PdfTextRun[] = value;
  if (!Array.isArray(input)) {
    if (input && typeof (input as unknown as { then?: unknown }).then === "function") void Promise.resolve(input).catch(() => {});
    throw new TypeError("PDF text must be a string or a synchronous array of text runs.");
  }
  if (input.length > settings.limits.maxTextRuns) throw new OfficeIMOError("RESOURCE_LIMIT", "maxTextRuns exceeded.");
  let characters = 0, budgetCharacters = 0;
  const runs: CapturedRun[] = [];
  for (const run of input) {
    if (!run || typeof run !== "object" || Object.keys(run).some(key => !["text", "bold", "italic", "color", "link", "font"].includes(key)))
      throw new TypeError("PDF text runs support text, bold, italic, color, link and font.");
    const { text, bold, italic, color, link, font } = run;
    if (typeof text !== "string") throw new TypeError("PDF run text must be a string.");
    if (font !== undefined) { if (!(font instanceof PdfFont)) throw new TypeError("A PDF run font must be a PdfFont."); fonts.registerFont(font); }
    characters += text.length;
    if (characters > settings.limits.maxCellCharacters) throw new OfficeIMOError("RESOURCE_LIMIT", "maxCellCharacters exceeded.");
    if (reserveBudget) budgetCharacters += pdfBudgetText(text, settings.budget).length;
    const style = presentation({ ...base, ...(bold === undefined ? {} : { bold }), ...(italic === undefined ? {} : { italic }), ...(color === undefined ? {} : { color }) });
    if (settings.options.tagged && link !== undefined && text.length && !text.trim()) throw new TypeError("Tagged PDF inline links require nonempty displayed text.");
    runs.push(Object.freeze({ text, style, ...(font === undefined ? {} : { font }), ...(link === undefined ? {} : { link: copyExportLink(link) }) }));
  }
  // Reserve before allocating a joined rich-text value or retaining it in a kept paragraph chain.
  if (reserveBudget && (characters || countEmpty)) settings.budget.reserve(1, budgetCharacters);
  return Object.freeze({ text: runs.map(run => run.text).join(""), runs: Object.freeze(runs) });
}

/** @internal Greedy wrapping over the continuous text stream, preserving run emphasis and URI geometry. */
export async function wrapPdfText(captured: PdfCapturedText, fonts: PdfFontResources, size: number, width: number,
  settings: PdfSettings, wrap = true): Promise<readonly PdfInlineLine[]> {
  if (settings.options.textShaper) return wrapShapedPdfText(captured, fonts, size, width, settings, wrap);
  const { signal } = settings.options, { text, runs } = captured;
  checkAbort(signal);
  // Keep font-selection order while deferring a single-run face array until wrapping needs it.
  const single = runs.length === 1 ? runs[0] : undefined;
  const singleFace = single && fonts.select(single.style.bold, single.style.italic, single.font);
  const selectedFaces = singleFace ? undefined : runs.map(run => fonts.select(run.style.bold, run.style.italic, run.font));
  const fallback = fonts.select();
  // Short fitting runs need no per-character layout arrays. Longer text keeps the existing
  // cooperative checkpoints; controls or the first overflowing advance use the general path.
  if (single && text.length > 0 && text.length < 4096) {
    const run = single, font = singleFace!;
    let widthSoFar = 0, fits = true;
    for (let offset = 0; offset < text.length;) {
      const scalar = text.codePointAt(offset)!;
      if (scalar === 9 || scalar === 10 || scalar === 13) { fits = false; break; }
      const advance = font.width(scalar) * size / 1000;
      if (advance > width + .001 || widthSoFar + advance > width + .001) { fits = false; break; }
      widthSoFar += advance; offset += scalar > 0xffff ? 2 : 1;
    }
    if (fits) {
      if (settings.limits.maxRowLines < 1) throw new OfficeIMOError("RESOURCE_LIMIT", "maxRowLines exceeded.");
      return [{ text, width: widthSoFar, segments: [{ text, width: widthSoFar, font,
        ...(run.style.color === undefined ? {} : { color: run.style.color }), ...(run.link ? { link: run.link } : {}) }],
        lineHeight: font.lineHeight(size), ascent: Math.max(0, (font.program?.ascent ?? 800) * size / 1000) }];
    }
  }
  const faces = selectedFaces ?? [singleFace!], lines: PdfInlineLine[] = [];
  let chars: string[] = [], advances: number[] = [], owners: number[] = [], measured = 0, breakAt = -1;
  let runIndex = 0, runEnd = runs[0]?.text.length ?? 0, scanned = 0;
  const emit = (count: number): void => {
    if (lines.length >= settings.limits.maxRowLines) throw new OfficeIMOError("RESOURCE_LIMIT", "maxRowLines exceeded.");
    const segments: PdfInlineSegment[] = [];
    let lineWidth = 0, lineHeight = 0, ascent = 0;
    for (let begin = 0; begin < count;) {
      const owner = owners[begin]!, run = runs[owner]!, font = faces[owner]!;
      let end = begin + 1; while (end < count && owners[end] === owner) end++;
      const segmentWidth = advances.slice(begin, end).reduce((total, advance) => total + advance, 0);
      segments.push({ text: chars.slice(begin, end).join(""), width: segmentWidth, font,
        ...(run.style.color === undefined ? {} : { color: run.style.color }), ...(run.link ? { link: run.link } : {}) });
      lineWidth += segmentWidth; lineHeight = Math.max(lineHeight, font.lineHeight(size));
      ascent = Math.max(ascent, (font.program?.ascent ?? 800) * size / 1000); begin = end;
    }
    // Empty physical lines remain observable line breaks without creating zero-area links.
    if (!segments.length) {
      const font = faces[0] ?? fallback;
      segments.push({ text: "", width: 0, font, ...(runs[0]?.style.color === undefined ? {} : { color: runs[0].style.color }) });
      lineHeight = font.lineHeight(size); ascent = (font.program?.ascent ?? 800) * size / 1000;
    }
    lines.push({ text: chars.slice(0, count).join(""), width: lineWidth, segments, lineHeight, ascent });
    chars = chars.slice(count); advances = advances.slice(count); owners = owners.slice(count);
    measured = advances.reduce((total, advance) => total + advance, 0); breakAt = -1;
    for (let i = 0; i < chars.length; i++) if (chars[i] === " ") breakAt = i;
  };
  const append = (character: string, owner: number): void => {
    const advance = faces[owner]!.width(character.codePointAt(0)!) * size / 1000;
    if (advance > width + .001) throw new RangeError("PDF column is too narrow for a glyph. Increase page/column width or reduce padding/font size.");
    if (measured + advance > width + .001) {
      if (!wrap) throw new RangeError("Unwrapped PDF cell exceeds its column width.");
      emit(breakAt >= 0 ? breakAt + 1 : chars.length);
      if (measured + advance > width + .001) emit(chars.length);
    }
    chars.push(character); advances.push(advance); owners.push(owner); measured += advance;
    if (character === " ") breakAt = chars.length - 1;
  };
  for (let offset = 0; offset < text.length;) {
    while (offset >= runEnd && runIndex < runs.length - 1) runEnd += runs[++runIndex]!.text.length;
    const owner = runIndex, scalar = text.codePointAt(offset)!, character = String.fromCodePoint(scalar);
    if (offset + character.length > runEnd && runIndex < runs.length - 1) {
      let next = runIndex + 1; while (!runs[next]!.text.length) next++;
      const first = runs[owner]!, other = runs[next]!, firstFace = faces[owner]!, nextFace = faces[next]!;
      const sameFace = firstFace.font === nextFace.font && (firstFace.font ?
        firstFace.syntheticBold === nextFace.syntheticBold && firstFace.syntheticItalic === nextFace.syntheticItalic :
        firstFace.bold === nextFace.bold && firstFace.italic === nextFace.italic);
      if (!sameFace || color(first.style.color ?? "111827") !== color(other.style.color ?? "111827") ||
        first.link?.target !== other.link?.target || first.link?.tooltip !== other.link?.tooltip)
        throw new NotSupportedError("A PDF presentation boundary cannot split a UTF-16 scalar.");
    }
    offset += character.length;
    if (character === "\r" || character === "\n") {
      if (character === "\r" && text[offset] === "\n") offset++;
      emit(chars.length);
    } else if (character === "\t") { for (let i = 0; i < 4; i++) append(" ", owner); }
    else append(character, owner);
    if ((++scanned & 4095) === 0) { if (taskYieldDue()) await pause(); checkAbort(signal); }
  }
  emit(chars.length);
  return lines;
}
