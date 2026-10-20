import type { CellPresentation } from "../core/index.js";
import type { PdfOptions } from "./types.js";
import type { PdfDrawingBlock } from "./drawing-types.js";
import type { PdfText } from "./inline-types.js";
import type { PdfBookmark, PdfSectionBlock } from "./section-types.js";

/** Page geometry, fonts, decorations and resource policy shared by every document block. */
export interface PdfDocumentOptions extends Omit<PdfOptions<never>, "columns" | "includeHeader" | "headerRows" | "footer" |
  "headerPresentation" | "footerPresentation" | "alternateRowColor" | "columnWidths" | "wideTable" | "formatValue" | "formatRuns"> {
  /** Named paragraph appearance and keep rules, copied before the block source is consumed. No style inheritance. */
  readonly paragraphStyles?: Readonly<Record<string, PdfParagraphStyle>>;
}
/** A table's own columns, repeated headings, presentation and numeric totals. */
export type PdfTableOptions<T = never> = Pick<PdfOptions<T>, "columns" | "includeHeader" | "headerRows" | "footer" |
  "headerPresentation" | "footerPresentation" | "alternateRowColor" | "columnWidths" | "wideTable" | "formatValue" | "formatRuns">;
/** Construct through pdfTable so column keys and getters are checked against the row type. */
export interface PdfTableBlock {
  readonly kind: "table";
  readonly rows: Iterable<object> | AsyncIterable<object>;
  readonly options: PdfTableOptions;
  /** Visible caption before the table, kept with its first heading/data fragment. */
  readonly caption?: PdfCaption;
  /** Repeated before leaf headings when this table continues on another page. Tagged output is unsupported. */
  readonly continuationCaption?: PdfCaption;
}
/** Reusable paragraph appearance, spacing and keep rules. Semantic roles remain on their content blocks. */
export interface PdfParagraphStyle {
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
/** Flowing text and inline emphasis/links, using the document's fonts and supported-script contract. */
export interface PdfParagraphBlock extends PdfParagraphStyle {
  readonly kind: "paragraph";
  readonly text: PdfText;
  /** Name from paragraphStyles. Explicit fields override it; presentation merges member by member. */
  readonly style?: string;
  /** Heading level in tagged output; omitted text is a paragraph. Appearance uses presentation. */
  readonly headingLevel?: 1 | 2 | 3 | 4 | 5 | 6;
  /** Native navigation destination at this paragraph's actual first painted line. */
  readonly bookmark?: PdfBookmark;
}
/** Visible label kept before its table or figure; numbering is authored as part of text. */
export interface PdfCaption extends Omit<PdfParagraphStyle, "keepTogether" | "keepWithNext"> {
  readonly text: PdfText;
  readonly style?: string;
}
/** One rich paragraph body. Explicit line breaks may separate a finding's title, message and scope. */
export interface PdfListItem {
  readonly text: PdfText;
  readonly style?: string;
  /** Move this whole item intact; reject when it exceeds a fresh page. */
  readonly keepTogether?: boolean;
}
/** A streamed semantic list, with hanging body text and one marker per item. Nested blocks are unsupported. */
export interface PdfListBlock {
  readonly kind: "list";
  readonly items: Iterable<PdfListItem> | AsyncIterable<PdfListItem>;
  /** Decimal numbering when true; otherwise a bullet. Default false. */
  readonly ordered?: boolean;
  /** First positive safe-integer number; valid only for ordered lists. Default one. */
  readonly start?: number;
  /** Default paragraph style for items; an item's style replaces this name. keepWithNext is unsupported. */
  readonly itemStyle?: string;
  /** Body offset in points; default reserves at least two ems and enough room for the marker and gap. */
  readonly indent?: number;
  /** Gap between marker and body in points; default half an em. */
  readonly markerGap?: number;
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
  /** Visible caption before the complete image; distinct from alternativeText. */
  readonly caption?: PdfCaption;
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
  readonly caption?: PdfCaption;
}
/** Blocks are consumed once in order; each table owns its row source until that table completes. */
export type PdfDocumentBlock = PdfParagraphBlock | PdfGroupBlock | PdfTableBlock | PdfPageBreakBlock | PdfImageBlock | PdfVectorBlock | PdfDrawingBlock | PdfListBlock | PdfSectionBlock;
