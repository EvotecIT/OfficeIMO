import { NotSupportedError, OfficeIMOError } from "../core/errors.js";
import { checkAbort, pause, taskYieldDue, withAbort } from "../core/iteration.js";
import type { PdfFontResource, PdfFontResources } from "./font-resources.js";
import type { PdfCapturedText, PdfInlineLine, PdfInlineSegment } from "./inline-text.js";
import type { PdfSettings } from "./settings.js";
import type { PdfShapedGlyph, PdfShapedRun, PdfShapingRun } from "./text-shaping-types.js";

interface Part { text: string; owners: number[]; }
interface Glyph extends PdfShapedGlyph { readonly logicalText: string; }
interface Validated { readonly runs: readonly PdfShapedRun[]; readonly boundaries: readonly number[]; readonly advances: ReadonlyMap<number, number>; readonly width: number; }
const boundary = (text: string, at: number): boolean => at === 0 || at === text.length || !(text.charCodeAt(at) >= 0xdc00 && text.charCodeAt(at) <= 0xdfff && text.charCodeAt(at - 1) >= 0xd800 && text.charCodeAt(at - 1) <= 0xdbff);

/** Normalize physical line breaks/tabs with the scalar writer's semantics, retaining presentation ownership. */
function parts(captured: PdfCapturedText, maxLines: number): readonly Part[] {
  const result: Part[] = [{ text: "", owners: [] }]; let owner = 0, end = captured.runs[0]?.text.length ?? 0;
  for (let offset = 0; offset < captured.text.length;) {
    while (offset >= end && owner < captured.runs.length - 1) end += captured.runs[++owner]!.text.length;
    const scalar = captured.text.codePointAt(offset)!, character = String.fromCodePoint(scalar); offset += character.length;
    if (scalar >= 0xd800 && scalar <= 0xdfff) throw new TypeError("PDF text contains an unpaired UTF-16 surrogate.");
    if (offset > end && owner < captured.runs.length - 1) throw new NotSupportedError("A shaped presentation boundary cannot split a UTF-16 scalar.");
    if (character === "\n" || character === "\r") {
      if (character === "\r" && captured.text[offset] === "\n") offset++;
      if (result.length >= maxLines) throw new OfficeIMOError("RESOURCE_LIMIT", "maxRowLines exceeded.");
      result.push({ text: "", owners: [] }); continue;
    }
    if (scalar < 32 && scalar !== 9 || scalar >= 127 && scalar < 160) throw new NotSupportedError("PDF text contains an unsupported control character.");
    const value = character === "\t" ? "    " : character, part = result.at(-1)!;
    if (part.text.length + value.length > 4096) throw new OfficeIMOError("RESOURCE_LIMIT", "Shaped PDF text supports at most 4,096 UTF-16 units per physical paragraph.");
    part.text += value; for (let i = 0; i < value.length; i++) part.owners.push(owner);
  }
  return result;
}

/** Validate a replaceable provider without trusting its font, glyph, metric or source-range claims. */
function validate(output: readonly PdfShapedRun[], part: Part, start: number, end: number, faces: readonly PdfFontResource[], captured: PdfCapturedText): Validated {
  if (!Array.isArray(output) || output.length > 32768) throw new TypeError("Invalid shaped run array.");
  let count = 0, width = 0;
  const covered = new Uint8Array(end - start), intervals = new Map<number, number>(), advances = new Map<number, number>();
  const copied: PdfShapedRun[] = [];
  for (const run of output) {
    if (!run || !Array.isArray(run.glyphs) || !faces.some(face => face.font === run.font)) throw new TypeError("Shaping must use an exact registered PdfFont identity.");
    count += run.glyphs.length; if (count > 32768) throw new OfficeIMOError("RESOURCE_LIMIT", "Shaped PDF text exceeded 32,768 glyphs.");
    const glyphs: PdfShapedGlyph[] = [];
    for (const glyph of run.glyphs) {
      const { id, clusterStart: a, clusterEnd: b, xAdvance, yAdvance, xOffset, yOffset } = glyph ?? {};
      const face = faces[part.owners[a] ?? -1];
      if (!face || face.font !== run.font || !face.program || !Number.isSafeInteger(id) || id <= 0 || id >= face.program.glyphCount)
        throw new TypeError("Shaping returned an invalid glyph or font identity.");
      if (!Number.isSafeInteger(a) || !Number.isSafeInteger(b) || a < start || b <= a || b > end || !boundary(part.text, a) || !boundary(part.text, b))
        throw new TypeError("Shaping returned an invalid UTF-16 cluster range.");
      if (![xAdvance, yAdvance, xOffset, yOffset].every(value => Number.isFinite(value) && Math.abs(value) <= 100000) || xAdvance < 0 || yAdvance !== 0)
        throw new TypeError("Shaping requires finite bounded horizontal glyph metrics.");
      const first = captured.runs[part.owners[a]!]!;
      for (let i = a; i < b; i++) {
        const other = captured.runs[part.owners[i]!]!;
        if (faces[part.owners[i]!]!.font !== run.font || first.style.bold !== other.style.bold || first.style.italic !== other.style.italic || first.style.color !== other.style.color || first.link !== other.link)
          throw new NotSupportedError("A PDF presentation, font or link boundary cannot split an indivisible shaped cluster.");
        covered[i - start] = 1;
      }
      const previous = intervals.get(a); if (previous !== undefined && previous !== b) throw new TypeError("Inconsistent shaped cluster ranges.");
      intervals.set(a, b); advances.set(a, (advances.get(a) ?? 0) + xAdvance); width += xAdvance;
      glyphs.push(Object.freeze({ id, clusterStart: a, clusterEnd: b, xAdvance, yAdvance, xOffset, yOffset }));
    }
    copied.push(Object.freeze({ font: run.font, glyphs: Object.freeze(glyphs) }));
  }
  if (covered.some(value => !value)) throw new TypeError("Shaping omitted logical source text.");
  const sorted = [...intervals].sort((a, b) => a[0] - b[0]); let cursor = start;
  for (const [a, b] of sorted) { if (a !== cursor) throw new TypeError("Shaped clusters overlap or leave a logical gap."); cursor = b; }
  return { runs: copied, boundaries: sorted.map(([, b]) => b), advances, width };
}

/** Core-owned bounded line flow. Shape finalized lines again so Arabic joining reflects the actual break. */
export async function wrapShapedPdfText(captured: PdfCapturedText, fonts: PdfFontResources, size: number, width: number,
  settings: PdfSettings, wrap: boolean): Promise<readonly PdfInlineLine[]> {
  const faces = captured.runs.map(run => fonts.select(run.style.bold, run.style.italic, run.font)), fallback = fonts.select();
  if (faces.some((face, i) => !face.font && captured.runs[i]!.text.length)) throw new NotSupportedError("Shaped PDF text requires explicit embedded TrueType fonts.");
  const lines: PdfInlineLine[] = [], provider = settings.options.textShaper!, signal = settings.options.signal;
  const emit = (part: Part, start: number, end: number, shaped?: Validated): void => {
    if (lines.length >= settings.limits.maxRowLines) throw new OfficeIMOError("RESOURCE_LIMIT", "maxRowLines exceeded.");
    const segments: PdfInlineSegment[] = []; let lineHeight = fallback.lineHeight(size), ascent = (fallback.program?.ascent ?? 800) * size / 1000;
    const seen = new Set<number>(); let owner = -1, segmentFont: PdfFontResource | undefined, glyphs: Glyph[] = [], segmentWidth = 0;
    const flush = (): void => {
      if (owner < 0 || !segmentFont) return; const run = captured.runs[owner]!;
      const logical = glyphs.filter(glyph => glyph.logicalText).sort((a, b) => a.clusterStart - b.clusterStart);
      const actualText = logical.map(glyph => glyph.logicalText).join("");
      segments.push({ text: actualText, actualText, logicalStart: Math.min(...glyphs.map(glyph => glyph.clusterStart)), width: segmentWidth * size / 1000, font: segmentFont, glyphs,
        ...(run.style.color === undefined ? {} : { color: run.style.color }), ...(run.link ? { link: run.link } : {}) });
      glyphs = []; segmentWidth = 0;
    };
    for (const run of shaped?.runs ?? []) for (const glyph of run.glyphs) {
      const nextOwner = part.owners[glyph.clusterStart]!, face = faces[nextOwner]!;
      if (owner !== nextOwner || segmentFont !== face) { flush(); owner = nextOwner; segmentFont = face; }
      const logicalText = seen.has(glyph.clusterStart) ? "" : part.text.slice(glyph.clusterStart, glyph.clusterEnd); seen.add(glyph.clusterStart);
      glyphs.push({ ...glyph, logicalText }); segmentWidth += glyph.xAdvance;
      lineHeight = Math.max(lineHeight, face.lineHeight(size)); ascent = Math.max(ascent, face.program!.ascent * size / 1000);
    }
    flush(); const text = part.text.slice(start, end);
    lines.push({ text, actualText: text, segments, width: (shaped?.width ?? 0) * size / 1000, lineHeight, ascent });
  };
  for (const part of parts(captured, settings.limits.maxRowLines)) {
    checkAbort(signal); if (!part.text) { emit(part, 0, 0); continue; }
    const runs: PdfShapingRun[] = [];
    for (let start = 0; start < part.text.length;) { const owner = part.owners[start]!; let end = start + 1; while (end < part.text.length && part.owners[end] === owner) end++;
      runs.push(Object.freeze({ start, end, font: faces[owner]!.font! })); start = end; }
    Object.freeze(runs);
    let calls = 0;
    const shape = async (start: number, end: number): Promise<Validated> => {
      if (++calls > 16385) throw new OfficeIMOError("RESOURCE_LIMIT", "Shaped line fitting exceeded its bounded work limit.");
      if (taskYieldDue()) await pause(); checkAbort(signal);
      const output = await withAbort(Promise.resolve(provider.shapeLine(Object.freeze({ text: part.text, start, end, runs,
        direction: settings.options.textDirection ?? "auto", ...(signal ? { signal } : {}) }))), signal);
      checkAbort(signal); return validate(output, part, start, end, faces, captured);
    };
    const initial = await shape(0, part.text.length), boundaries = initial.boundaries;
    if (!wrap) { if (initial.width * size / 1000 > width + .001) throw new RangeError("Unwrapped PDF cell exceeds its column width."); emit(part, 0, part.text.length, initial); continue; }
    let from = 0;
    while (from < boundaries.length) {
      const start = from ? boundaries[from - 1]! : 0; let to = from, measured = 0, space = -1;
      while (to < boundaries.length) { const at = to ? boundaries[to - 1]! : 0, advance = initial.advances.get(at)! * size / 1000;
        if (measured + advance > width + .001 && to > from) break;
        measured += advance; if (part.text[boundaries[to]! - 1] === " ") space = to + 1; to++;
        if (measured > width + .001) break;
      }
      if (to < boundaries.length && space > from) to = space;
      let final = start === 0 && to === boundaries.length ? initial : await shape(start, boundaries[to - 1]!);
      if (final.width * size / 1000 > width + .001) {
        let low = from + 1, high = to - 1, best: Validated | undefined, bestEnd = 0;
        while (low <= high) { const mid = Math.floor((low + high) / 2), candidate = await shape(start, boundaries[mid - 1]!);
          if (candidate.width * size / 1000 <= width + .001) { best = candidate; bestEnd = mid; low = mid + 1; } else high = mid - 1; }
        if (!best) throw new RangeError("PDF column is too narrow for an indivisible shaped cluster."); final = best; to = bestEnd;
      }
      emit(part, start, boundaries[to - 1]!, final); from = to;
    }
  }
  return lines;
}
