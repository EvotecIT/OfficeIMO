import { OfficeIMOError } from "../core/errors.js";
import type { CellPresentation } from "../core/index.js";
import type { PdfFontResources } from "./font-resources.js";
import type { PdfCaption, PdfParagraphBlock } from "./document-types.js";
import { capturePdfText, wrapPdfText } from "./inline-text.js";
import type { PdfCapturedText, PdfInlineLine } from "./inline-text.js";
import type { PdfSettings } from "./settings.js";
import { positive } from "./settings.js";
import { resolveParagraphStyle } from "./paragraph-styles.js";
import { captureBookmark } from "./section-types.js";
import type { PdfBookmark } from "./section-types.js";
import { captureFields } from "../internal/capture.js";
import { pdfBudgetText } from "./text.js";

export interface PdfCapturedParagraph {
  readonly text: PdfCapturedText;
  readonly style: Readonly<CellPresentation>;
  readonly size: number;
  readonly headingLevel?: number;
  readonly bookmark?: Readonly<PdfBookmark>;
  readonly keepTogether: boolean;
  readonly keepWithNext: boolean;
  readonly spaceBefore: number;
  readonly spaceAfter: number;
  /** Captions are bounded definitions; charge only their actual physical occurrences. */
  readonly budgetOnPaint?: boolean;
}
export interface PdfPreparedParagraph extends PdfCapturedParagraph {
  readonly lines: readonly PdfInlineLine[];
  readonly height: number;
  readonly firstHeight: number;
}
/** @internal Copy public paragraph input before any cooperative layout or source enumeration. */
export function captureParagraph(block: PdfParagraphBlock, settings: PdfSettings, fonts: PdfFontResources, budgetOnPaint = false): PdfCapturedParagraph {
  if (!block || block.kind !== "paragraph") throw new TypeError("PDF paragraph groups accept paragraph blocks only.");
  const { text, headingLevel, presentation: inputStyle, fontSize, keepTogether, keepWithNext, spaceBefore, spaceAfter, style: styleName, bookmark: inputBookmark } = block;
  if (headingLevel !== undefined && (!Number.isInteger(headingLevel) || headingLevel < 1 || headingLevel > 6))
    throw new TypeError("Paragraph headingLevel must be an integer from 1 through 6.");
  const resolved = resolveParagraphStyle(styleName, {
    ...(fontSize === undefined ? {} : { fontSize }), ...(inputStyle === undefined ? {} : { presentation: inputStyle }),
    ...(keepTogether === undefined ? {} : { keepTogether }), ...(keepWithNext === undefined ? {} : { keepWithNext }),
    ...(spaceBefore === undefined ? {} : { spaceBefore }), ...(spaceAfter === undefined ? {} : { spaceAfter })
  }, settings.paragraphStyles);
  const style = resolved.presentation ?? Object.freeze({}), size = positive(resolved.fontSize ?? settings.fontSize, "Paragraph font size", 144);
  const bookmark = captureBookmark(inputBookmark, settings);
  return Object.freeze({ text: capturePdfText(text, style, settings, fonts, false, !budgetOnPaint), style, size,
    ...(budgetOnPaint ? { budgetOnPaint: true } : {}),
    ...(headingLevel === undefined ? {} : { headingLevel }), ...(bookmark === undefined ? {} : { bookmark }),
    keepTogether: resolved.keepTogether ?? false, keepWithNext: resolved.keepWithNext ?? false,
    spaceBefore: resolved.spaceBefore ?? 0, spaceAfter: resolved.spaceAfter ?? 8 });
}
/** @internal Captions reuse rich text layout while their actual table/figure supplies the semantic owner. */
export function captureCaption(caption: PdfCaption, settings: PdfSettings, fonts: PdfFontResources): PdfCapturedParagraph {
  if (!caption || typeof caption !== "object" || Array.isArray(caption) ||
    Object.keys(caption).some(key => !["text", "style", "fontSize", "presentation", "spaceBefore", "spaceAfter"].includes(key)))
    throw new TypeError("PDF captions support text, style, fontSize, presentation, spaceBefore and spaceAfter.");
  return captureParagraph({ ...captureFields(caption, ["text", "style", "fontSize", "presentation", "spaceBefore", "spaceAfter"]),
    kind: "paragraph", keepTogether: true, keepWithNext: true }, settings, fonts, true);
}
/** @internal Layout is bounded by the same per-cell text and wrapped-line limits as table cells. */
export async function prepareParagraph(captured: PdfCapturedParagraph, settings: PdfSettings, fonts: PdfFontResources,
  width = settings.page.width - settings.margins.left - settings.margins.right): Promise<PdfPreparedParagraph> {
  if (!captured.text.text) return { ...captured, lines: [], height: 0, firstHeight: 0 };
  if (!captured.budgetOnPaint) {
    const budgetText = pdfBudgetText(captured.text.text, settings.budget);
    settings.budget.cell(budgetText, budgetText.length);
  }
  const lines = await wrapPdfText(captured.text, fonts, captured.size, positive(width, "Paragraph width"), settings);
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
