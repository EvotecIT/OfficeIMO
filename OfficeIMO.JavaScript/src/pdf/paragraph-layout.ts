import { OfficeIMOError } from "../core/errors.js";
import type { CellPresentation } from "../core/index.js";
import type { PdfFontResources } from "./font-resources.js";
import type { PdfParagraphBlock } from "./document-types.js";
import { capturePdfText, wrapPdfText } from "./inline-text.js";
import type { PdfCapturedText, PdfInlineLine } from "./inline-text.js";
import type { PdfSettings } from "./settings.js";
import { positive, presentation } from "./settings.js";

export interface PdfCapturedParagraph {
  readonly text: PdfCapturedText;
  readonly style: Readonly<CellPresentation>;
  readonly size: number;
  readonly headingLevel?: number;
  readonly keepTogether: boolean;
  readonly keepWithNext: boolean;
  readonly spaceBefore: number;
  readonly spaceAfter: number;
}
export interface PdfPreparedParagraph extends PdfCapturedParagraph {
  readonly lines: readonly PdfInlineLine[];
  readonly height: number;
  readonly firstHeight: number;
}
/** @internal Copy public paragraph input before any cooperative layout or source enumeration. */
export function captureParagraph(block: PdfParagraphBlock, settings: PdfSettings, fonts: PdfFontResources): PdfCapturedParagraph {
  if (!block || block.kind !== "paragraph") throw new TypeError("PDF paragraph groups accept paragraph blocks only.");
  const { text, headingLevel, presentation: inputStyle, fontSize, keepTogether, keepWithNext, spaceBefore, spaceAfter } = block;
  if (headingLevel !== undefined && (!Number.isInteger(headingLevel) || headingLevel < 1 || headingLevel > 6))
    throw new TypeError("Paragraph headingLevel must be an integer from 1 through 6.");
  if (inputStyle && Object.keys(inputStyle).some(key => !["bold", "italic", "color", "alignment"].includes(key)))
    throw new TypeError("PDF paragraph presentation supports bold, italic, color and alignment.");
  for (const [name, flag] of [["keepTogether", keepTogether], ["keepWithNext", keepWithNext]] as const)
    if (flag !== undefined && typeof flag !== "boolean") throw new TypeError("Paragraph " + name + " must be boolean.");
  const style = presentation(inputStyle), size = positive(fontSize ?? settings.fontSize, "Paragraph font size", 144);
  const spacing = (value: number | undefined, fallback: number): number => {
    if (value === undefined) return fallback;
    if (!Number.isFinite(value) || value < 0 || value > 14400) throw new RangeError("Paragraph spacing must be from zero through 14,400 points.");
    return value;
  };
  return Object.freeze({ text: capturePdfText(text, style, settings, fonts, false), style, size,
    ...(headingLevel === undefined ? {} : { headingLevel }), keepTogether: keepTogether ?? false, keepWithNext: keepWithNext ?? false,
    spaceBefore: spacing(spaceBefore, 0), spaceAfter: spacing(spaceAfter, 8) });
}
/** @internal Layout is bounded by the same per-cell text and wrapped-line limits as table cells. */
export async function prepareParagraph(captured: PdfCapturedParagraph, settings: PdfSettings, fonts: PdfFontResources): Promise<PdfPreparedParagraph> {
  if (!captured.text.text) return { ...captured, lines: [], height: 0, firstHeight: 0 };
  settings.budget.cell(captured.text.text, captured.text.text.length);
  const lines = await wrapPdfText(captured.text, fonts, captured.size, settings.page.width - settings.margins.left - settings.margins.right, settings);
  return { ...captured, lines, height: captured.spaceBefore + lines.reduce((height, line) => height + line.lineHeight, 0) + captured.spaceAfter,
    firstHeight: captured.spaceBefore + lines[0]!.lineHeight };
}
/** @internal Pending keep chains/groups cannot retain an unbounded report or consume a table source. */
export function checkParagraphBuffer(paragraphs: readonly PdfCapturedParagraph[], settings: PdfSettings): void {
  if (paragraphs.length > settings.limits.maxGroupBlocks) throw new OfficeIMOError("RESOURCE_LIMIT", "maxGroupBlocks exceeded.");
  let characters = 0, runs = 0;
  for (const paragraph of paragraphs) {
    characters += paragraph.text.text.length; runs += paragraph.text.runs.length;
    if (characters > settings.limits.maxCellCharacters) throw new OfficeIMOError("RESOURCE_LIMIT", "maxCellCharacters exceeded by a paragraph group or keep chain.");
    if (runs > settings.limits.maxTextRuns) throw new OfficeIMOError("RESOURCE_LIMIT", "maxTextRuns exceeded by a paragraph group or keep chain.");
  }
}
/** @internal Minimum leading fragment compatible with the first visible paragraph's own keep requests. */
export function paragraphStartHeight(paragraphs: readonly PdfPreparedParagraph[]): number {
  const visible = paragraphs.filter(paragraph => paragraph.lines.length > 0);
  let required = 0;
  for (let i = visible.length - 1; i >= 0; i--) {
    const paragraph = visible[i]!;
    required = paragraph.keepWithNext ? paragraph.height + required : paragraph.keepTogether ? paragraph.height : paragraph.firstHeight;
  }
  return required;
}
