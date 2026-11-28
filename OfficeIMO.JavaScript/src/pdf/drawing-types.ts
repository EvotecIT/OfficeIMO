import type { PdfCaption, PdfPathCommand } from "./document-types.js";

/** Versioned portable drawing. Lengths are points, with a top-left origin and downward y axis. */
export interface PdfDrawing {
  readonly version: 1;
  readonly width: number;
  readonly height: number;
  readonly elements: readonly PdfDrawingElement[];
  /** Paint text whose slot has no supplied TrueType face with the standard Helvetica face (WinAnsi characters only)
   * instead of rejecting it. Default false: drawing text requires exact embedded faces. */
  readonly standardFonts?: boolean;
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
  /** 1 through 16 stroke dash lengths from 0 through 14,400 points. At least one must survive rounding to 0.001 pt. */
  readonly dash?: readonly number[];
  /** Dash phase from -1,000,000 through 1,000,000 points; requires dash. */
  readonly dashOffset?: number;
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
  /** Which point of the measured run sits at x and y. Default start. */
  readonly anchor?: "start" | "middle" | "end";
  /** Clockwise baseline rotation in degrees about the anchor, -360 through 360. Default 0. */
  readonly rotation?: number;
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
  readonly caption?: PdfCaption;
}
