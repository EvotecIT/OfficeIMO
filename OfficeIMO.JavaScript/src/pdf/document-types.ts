import type { CellPresentation } from "../core/index.js";
import type { PdfOptions } from "./types.js";
import type { PdfDrawingBlock } from "./drawing-types.js";
import type { PdfText } from "./inline-types.js";

/** Page geometry, fonts, decorations and resource policy shared by every document block. */
export interface PdfDocumentOptions extends Omit<PdfOptions<never>, "columns" | "includeHeader" | "headerRows" | "footer" |
  "headerPresentation" | "footerPresentation" | "alternateRowColor" | "columnWidths" | "wideTable" | "formatValue" | "formatRuns"> {}
/** A table's own columns, repeated headings, presentation and numeric totals. */
export type PdfTableOptions<T = never> = Pick<PdfOptions<T>, "columns" | "includeHeader" | "headerRows" | "footer" |
  "headerPresentation" | "footerPresentation" | "alternateRowColor" | "columnWidths" | "wideTable" | "formatValue" | "formatRuns">;
/** Construct through pdfTable so column keys and getters are checked against the row type. */
export interface PdfTableBlock {
  readonly kind: "table";
  readonly rows: Iterable<object> | AsyncIterable<object>;
  readonly options: PdfTableOptions;
}
/** Flowing text and inline emphasis/links, using the document's fonts and supported-script contract. */
export interface PdfParagraphBlock {
  readonly kind: "paragraph";
  readonly text: PdfText;
  /** Heading level in tagged output; omitted text is a paragraph. Appearance uses presentation. */
  readonly headingLevel?: 1 | 2 | 3 | 4 | 5 | 6;
  readonly fontSize?: number;
  readonly presentation?: Pick<CellPresentation, "bold" | "italic" | "color" | "alignment">;
  /** Move this paragraph intact; reject when its complete layout exceeds a fresh page. */
  readonly keepTogether?: boolean;
  /** Keep this paragraph with the next paragraph line, table heading/data fragment or complete graphic. */
  readonly keepWithNext?: boolean;
  /** Space in points, before this paragraph; default zero. */
  readonly spaceBefore?: number;
  /** Space in points, after this paragraph; default eight. */
  readonly spaceAfter?: number;
}
/** An explicitly bounded group of paragraphs. Keeps stay within the group; other block kinds and nested groups are rejected. */
export interface PdfGroupBlock {
  readonly kind: "group";
  readonly blocks: readonly PdfParagraphBlock[];
  /** Keep the complete paragraph group on one page; reject an oversized group. Default true. */
  readonly keepTogether?: boolean;
}
/** Explicitly finish the current page; the next block starts on a new page. */
export interface PdfPageBreakBlock { readonly kind: "pageBreak"; }
/** Already-decoded row-major 8-bit pixels, with no row padding or premultiplied alpha. */
export interface PdfRawImage {
  readonly format: "rgb" | "rgba";
  readonly width: number;
  readonly height: number;
  readonly data: Uint8Array;
}
/** An 8-bit grayscale/RGB JPEG. CMYK, ICC-profiled and nonstandard component layouts are rejected. */
export interface PdfJpegImage { readonly format: "jpeg"; readonly data: Uint8Array; }
/** A complete static unprofiled, noninterlaced 8-bit RGB/RGBA PNG, decoded with native DecompressionStream. */
export interface PdfPngImage { readonly format: "png"; readonly data: Uint8Array; }
/** An indivisible image. Oversized display dimensions scale down to the printable page. */
export interface PdfImageBlock {
  readonly kind: "image";
  readonly image: PdfRawImage | PdfJpegImage | PdfPngImage;
  /** Display lengths in PDF points; omitted dimensions retain the intrinsic aspect ratio at 96 DPI. */
  readonly width?: number;
  readonly height?: number;
  readonly alignment?: "left" | "center" | "right";
  /** Required for meaningful figures in tagged output; cannot accompany decorative:true. */
  readonly alternativeText?: string;
  /** Explicitly exclude a purely decorative graphic from the logical reading order. */
  readonly decorative?: boolean;
}
/** Coordinates use points from the vector block's top-left corner. */
export type PdfPathCommand = { readonly op: "move" | "line"; readonly x: number; readonly y: number } |
  { readonly op: "cubic"; readonly x1: number; readonly y1: number; readonly x2: number; readonly y2: number; readonly x: number; readonly y: number } |
  { readonly op: "close" };
/** A bounded path, painted once. No SVG, CSS, gradients, text or external resources are interpreted. */
export interface PdfVectorBlock {
  readonly kind: "vector";
  readonly width: number;
  readonly height: number;
  readonly commands: readonly PdfPathCommand[];
  readonly fill?: string;
  readonly stroke?: string;
  readonly strokeWidth?: number;
  readonly fillRule?: "nonzero" | "evenodd";
  readonly alignment?: "left" | "center" | "right";
  readonly alternativeText?: string;
  readonly decorative?: boolean;
}
/** Blocks are consumed once in order; each table owns its row source until that table completes. */
export type PdfDocumentBlock = PdfParagraphBlock | PdfGroupBlock | PdfTableBlock | PdfPageBreakBlock | PdfImageBlock | PdfVectorBlock | PdfDrawingBlock;
