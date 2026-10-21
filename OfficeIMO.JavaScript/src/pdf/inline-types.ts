import type { ExportLink } from "../core/links.js";
import type { PdfFont } from "./font.js";

/** One text run. Omitted emphasis/color inherits the paragraph or cell presentation. */
export interface PdfTextRun {
  readonly text: string;
  readonly bold?: boolean;
  readonly italic?: boolean;
  readonly color?: string;
  /** Exact embedded face. Bold/italic flags apply synthetic emphasis to an explicit face. No font fallback occurs. */
  readonly font?: PdfFont;
  /** Native URI annotation over each painted run fragment. */
  readonly link?: ExportLink;
}

/** Literal text or ordered inline runs; run boundaries do not introduce spaces or line breaks. */
export type PdfText = string | readonly PdfTextRun[];
