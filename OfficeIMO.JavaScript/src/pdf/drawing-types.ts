import type { PdfPathCommand } from "./document-types.js";

/** Versioned portable drawing. Lengths are points, with a top-left origin and downward y axis. */
export interface PdfDrawing {
  readonly version: 1;
  readonly width: number;
  readonly height: number;
  readonly elements: readonly PdfDrawingElement[];
}
/** One painted path; colors are six-digit RGB hex strings. */
export interface PdfDrawingPath {
  readonly kind: "path";
  readonly commands: readonly PdfPathCommand[];
  readonly fill?: string;
  readonly stroke?: string;
  readonly strokeWidth?: number;
  readonly fillRule?: "nonzero" | "evenodd";
  readonly lineCap?: "butt" | "round" | "square";
  readonly lineJoin?: "miter" | "round" | "bevel";
}
/** A face explicitly supplied in the PDF document's embedded static TrueType family. */
export type PdfDrawingFont = "regular" | "bold" | "italic" | "boldItalic";
/** Fixed baseline text. No wrapping, shaping, font lookup, synthetic emphasis or line breaks. */
export interface PdfDrawingText {
  readonly kind: "text";
  readonly text: string;
  readonly x: number;
  readonly y: number;
  readonly font: PdfDrawingFont;
  readonly fontSize: number;
  readonly color: string;
}
export interface PdfDrawingClip { readonly x: number; readonly y: number; readonly width: number; readonly height: number; }
/** Ordered children, optionally restricted by an axis-aligned rectangular clip. */
export interface PdfDrawingGroup {
  readonly kind: "group";
  readonly clip?: PdfDrawingClip;
  readonly elements: readonly PdfDrawingElement[];
}
export type PdfDrawingElement = PdfDrawingPath | PdfDrawingText | PdfDrawingGroup;
/** An indivisible drawing, automatically scaled down to fit the printable page. */
export interface PdfDrawingBlock {
  readonly kind: "drawing";
  readonly drawing: PdfDrawing;
  readonly alignment?: "left" | "center" | "right";
  readonly alternativeText?: string;
  readonly decorative?: boolean;
}
