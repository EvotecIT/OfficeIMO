import type { PdfFont } from "./font.js";

/** Horizontal text direction. Auto resolves a base direction for each physical paragraph. */
export type PdfTextDirection = "auto" | "ltr" | "rtl";
/** A registered font covering a half-open UTF-16 range of the physical paragraph. */
export interface PdfShapingRun { readonly start: number; readonly end: number; readonly font: PdfFont; }
/** One positioned glyph. Metrics use 1,000 units per em; clusters are logical, half-open UTF-16 ranges. */
export interface PdfShapedGlyph {
  readonly id: number;
  readonly clusterStart: number;
  readonly clusterEnd: number;
  readonly xAdvance: number;
  readonly yAdvance: number;
  readonly xOffset: number;
  readonly yOffset: number;
}
/** Glyphs and runs are in left-to-right painting order, regardless of their logical text direction. */
export interface PdfShapedRun { readonly font: PdfFont; readonly glyphs: readonly PdfShapedGlyph[]; }
/** The entire paragraph is supplied so a wrapped line retains paragraph bidi context. */
export interface PdfShapeLineRequest {
  readonly text: string;
  readonly start: number;
  readonly end: number;
  readonly runs: readonly PdfShapingRun[];
  readonly direction: PdfTextDirection;
  readonly signal?: AbortSignal;
}
/** Replaceable shaping boundary. The PDF writer validates output and owns line flow, fonts and PDF objects.
 * Requests contain at most 4,096 UTF-16 units. Return at most 32,768 horizontal glyphs.
 * Providers must preserve logical clusters and must not silently substitute missing fonts or glyphs. */
export interface PdfTextShaper {
  shapeLine(request: PdfShapeLineRequest): PromiseLike<readonly PdfShapedRun[]> | readonly PdfShapedRun[];
}
