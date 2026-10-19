import type { CellPresentation } from "../core/index.js";
import type { PdfOptions } from "./types.js";

/** Page geometry, fonts, decorations and resource policy shared by every document block. */
export interface PdfDocumentOptions extends Omit<PdfOptions<never>, "columns" | "includeHeader" | "headerRows" | "footer" |
  "headerPresentation" | "footerPresentation" | "alternateRowColor" | "columnWidths" | "wideTable" | "formatValue"> {}
/** A table's own columns, repeated headings, presentation and numeric totals. */
export type PdfTableOptions<T = never> = Pick<PdfOptions<T>, "columns" | "includeHeader" | "headerRows" | "footer" |
  "headerPresentation" | "footerPresentation" | "alternateRowColor" | "columnWidths" | "wideTable" | "formatValue">;
/** Construct through pdfTable so column keys and getters are checked against the row type. */
export interface PdfTableBlock {
  readonly kind: "table";
  readonly rows: Iterable<object> | AsyncIterable<object>;
  readonly options: PdfTableOptions;
}
/** Flowing scalar text, using the document's fonts and supported-script contract. */
export interface PdfParagraphBlock {
  readonly kind: "paragraph";
  readonly text: string;
  readonly fontSize?: number;
  readonly presentation?: Pick<CellPresentation, "bold" | "italic" | "color" | "alignment">;
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
/** An indivisible image. Oversized display dimensions scale down to the printable page. */
export interface PdfImageBlock {
  readonly kind: "image";
  readonly image: PdfRawImage | PdfJpegImage;
  /** Display lengths in PDF points; omitted dimensions retain the intrinsic aspect ratio at 96 DPI. */
  readonly width?: number;
  readonly height?: number;
  readonly alignment?: "left" | "center" | "right";
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
}
/** Blocks are consumed once in order; each table owns its row source until that table completes. */
export type PdfDocumentBlock = PdfParagraphBlock | PdfTableBlock | PdfPageBreakBlock | PdfImageBlock | PdfVectorBlock;
