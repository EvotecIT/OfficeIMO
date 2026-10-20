import { checkAbort } from "../../core/iteration.js";
import { NotSupportedError, OfficeIMOError } from "../../core/errors.js";
import type { PdfFont } from "../font.js";
import type { PdfShapeLineRequest, PdfShapedGlyph, PdfShapedRun, PdfTextShaper } from "../text-shaping-types.js";
import { isUnicode13ProfileScalar } from "./unicode13-profile.js";

interface HbFont { setScale(x: number, y: number): void; }
interface HbBuffer {
  add(codePoint: number, cluster: number): void;
  reset(): void;
  guessSegmentProperties(): void;
  setDirection(direction: 4 | 5): void;
  setLanguage(language: string): void;
  setClusterLevel(level: 0): void;
  getGlyphInfos(): { codepoint: number; cluster: number }[];
  getGlyphPositions(): { xAdvance: number; yAdvance: number; xOffset: number; yOffset: number }[];
  getLength(): number;
}
/** Structural subset of harfbuzzjs 1.6.3. The host imports its ESM/WASM module explicitly. */
export interface PdfHarfBuzzModule {
  readonly Blob: new(bytes: ArrayBuffer) => object;
  readonly Face: new(blob: never) => object;
  readonly Font: new(face: never) => HbFont;
  readonly Buffer: new() => HbBuffer;
  shape(font: never, buffer: never): void;
}
export interface PdfBidiLevels { readonly levels: Uint8Array; readonly paragraphs: readonly { start: number; end: number; level: number }[]; }
/** Structural subset of an initialized bidi-js 1.1.0 instance (Unicode 13.0 data). */
export interface PdfBidiEngine {
  getEmbeddingLevels(text: string, direction?: "auto" | "ltr" | "rtl"): PdfBidiLevels;
  getReorderedIndices(text: string, levels: PdfBidiLevels, start?: number, end?: number): number[];
  getBidiCharTypeName(character: string): string;
}
export interface PdfHarfBuzzOptions {
  readonly harfbuzz: PdfHarfBuzzModule;
  readonly bidi: PdfBidiEngine;
  /** Optional BCP 47 shaping language; defaults to und. */
  readonly language?: string;
}
export interface PdfHarfBuzzTextShaper extends PdfTextShaper {
  /** Drop strong font/buffer references and reject future calls. HarfBuzz v1 frees native objects through finalizers. */
  dispose(): void;
}
interface Item { start: number; end: number; font: PdfFont; level: number; }

/** Adapt host-loaded HarfBuzz and bidi engines. No fetch, package import, font fallback or WASM initialization occurs here.
 * Qualified profile: horizontal static TrueType fonts, at most 16 cached font identities, 4,096 UTF-16 units per paragraph.
 * Cancellation is checked between bounded synchronous engine calls; worker termination owns deterministic WASM lifetime. */
export function createHarfBuzzTextShaper(options: PdfHarfBuzzOptions): PdfHarfBuzzTextShaper {
  const { harfbuzz: hb, bidi } = options;
  if (!hb || !bidi || [hb.Blob, hb.Face, hb.Font, hb.Buffer, hb.shape, bidi.getEmbeddingLevels, bidi.getReorderedIndices, bidi.getBidiCharTypeName].some(method => typeof method !== "function"))
    throw new TypeError("Supply initialized HarfBuzz and bidi engines.");
  const language = options.language ?? "und";
  if (typeof language !== "string" || language.length > 63 || Intl.getCanonicalLocales(language).length !== 1) throw new TypeError("Invalid shaping language.");
  const fonts = new Map<PdfFont, { blob: object; face: object; font: HbFont }>();
  let buffer: HbBuffer | undefined, disposed = false;
  let previousText: string | undefined, previousDirection: string | undefined, previousLevels: PdfBidiLevels | undefined;
  const getFont = (source: PdfFont): HbFont => {
    const cached = fonts.get(source); if (cached) return cached.font;
    if (fonts.size >= 16) throw new OfficeIMOError("RESOURCE_LIMIT", "The optional text provider supports at most 16 cached font identities.");
    // Constructor arguments are opaque because the loaded module owns its native handle classes.
    const bytes = source.toBytes(), blob = new hb.Blob(bytes.buffer as ArrayBuffer), face = new hb.Face(blob as never), font = new hb.Font(face as never);
    font.setScale(1000, 1000); fonts.set(source, { blob, face, font }); return font;
  };
  return {
    shapeLine(request: PdfShapeLineRequest): readonly PdfShapedRun[] {
      if (disposed) throw new TypeError("The text provider has been disposed.");
      checkAbort(request.signal);
      const { text, start, end, runs, direction } = request;
      if (!["auto", "ltr", "rtl"].includes(direction)) throw new TypeError("Invalid shaping direction.");
      if (text.length > 4096) throw new OfficeIMOError("RESOURCE_LIMIT", "The optional text provider supports at most 4,096 UTF-16 units per physical paragraph.");
      if (!Number.isSafeInteger(start) || !Number.isSafeInteger(end) || start < 0 || end < start || end > text.length) throw new TypeError("Invalid shaping line range.");
      if (start === end) return [];
      const levels = previousText === text && previousDirection === direction ? previousLevels! : bidi.getEmbeddingLevels(text, direction);
      previousText = text; previousDirection = direction; previousLevels = levels;
      if (!(levels.levels instanceof Uint8Array) || levels.levels.length !== text.length) throw new TypeError("Invalid bidi embedding levels.");
      const base = levels.paragraphs.find(paragraph => paragraph.start <= start && paragraph.end >= end - 1)?.level;
      if (base === undefined) throw new TypeError("A shaped line must stay within one physical bidi paragraph.");
      const lineText = text.slice(start, end), lineLevels = levels.levels.slice(start, end);
      // Resolve L1 trailing whitespace on the line-local view while retaining paragraph embedding levels.
      for (let i = lineText.length - 1; i >= 0 && ["S", "WS", "B", "LRI", "RLI", "FSI", "PDI", "BN", "RLE", "LRE", "RLO", "LRO", "PDF"].includes(bidi.getBidiCharTypeName(lineText[i]!)); i--) lineLevels[i] = base;
      const items: Item[] = [];
      for (let offset = start; offset < end;) {
        const run = runs.find(run => run.start <= offset && run.end > offset);
        if (!run) throw new TypeError("Every shaped character needs a registered font.");
        const scalar = text.codePointAt(offset)!, length = scalar > 0xffff ? 2 : 1, level = lineLevels[offset - start]!;
        if (scalar >= 0xd800 && scalar <= 0xdfff || offset + length > end) throw new TypeError("Shaping ranges must not split UTF-16 scalars.");
        if (!isUnicode13ProfileScalar(scalar))
          throw new NotSupportedError("The optional provider's Unicode 13.0 profile supports Latin, Arabic, Common and Inherited characters assigned in that version.");
        const previous = items.at(-1);
        // Same-font presentation boundaries remain in one shaping run, preserving contextual joining.
        if (previous && previous.font === run.font && previous.level === level) previous.end = offset + length;
        else items.push({ start: offset, end: offset + length, font: run.font, level });
        offset += length;
      }
      const visual = bidi.getReorderedIndices(lineText, { levels: lineLevels, paragraphs: [{ start: 0, end: lineText.length - 1, level: base }] }), ranks = new Map<number, number>();
      visual.forEach((offset, rank) => ranks.set(offset + start, rank));
      items.sort((a, b) => Math.min(...Array.from({ length: a.end - a.start }, (_, i) => ranks.get(a.start + i) ?? Infinity)) -
        Math.min(...Array.from({ length: b.end - b.start }, (_, i) => ranks.get(b.start + i) ?? Infinity)));
      const output: PdfShapedRun[] = []; let count = 0;
      for (const item of items) {
        checkAbort(request.signal); const font = getFont(item.font);
        buffer ??= new hb.Buffer(); buffer.reset();
        for (let offset = item.start; offset < item.end;) { const scalar = text.codePointAt(offset)!; buffer.add(scalar, offset); offset += scalar > 0xffff ? 2 : 1; }
        buffer.setClusterLevel(0); buffer.setDirection(item.level & 1 ? 5 : 4); buffer.setLanguage(language); buffer.guessSegmentProperties();
        hb.shape(font as never, buffer as never); checkAbort(request.signal);
        count += buffer.getLength(); if (count > 32768) throw new OfficeIMOError("RESOURCE_LIMIT", "The optional text provider exceeded 32,768 glyphs.");
        const infos = buffer.getGlyphInfos(), positions = buffer.getGlyphPositions();
        if (infos.length !== positions.length) throw new TypeError("HarfBuzz returned inconsistent glyph arrays.");
        const clusters = [...new Set(infos.map(info => info.cluster))].sort((a, b) => a - b), ends = new Map(clusters.map((cluster, i) => [cluster, clusters[i + 1] ?? item.end]));
        const glyphs: PdfShapedGlyph[] = infos.map((info, i) => ({ id: info.codepoint, clusterStart: info.cluster, clusterEnd: ends.get(info.cluster)!, ...positions[i]! }));
        if (glyphs.some(glyph => !glyph.id)) throw new NotSupportedError("The supplied font has a missing glyph in shaped text.");
        output.push({ font: item.font, glyphs });
      }
      return output;
    },
    dispose(): void { disposed = true; fonts.clear(); buffer = undefined; previousText = undefined; previousLevels = undefined; }
  };
}
