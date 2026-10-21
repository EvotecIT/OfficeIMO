import type { PdfListBlock, PdfListItem } from "./document-types.js";
import type { PdfFontResources } from "./font-resources.js";
import { capturePdfText, wrapPdfText } from "./inline-text.js";
import type { PdfInlineLine } from "./inline-text.js";
import { captureParagraph, prepareParagraph } from "./paragraph-layout.js";
import type { PdfPreparedParagraph } from "./paragraph-layout.js";
import { resolveParagraphStyle } from "./paragraph-styles.js";
import type { PdfSettings } from "./settings.js";

/** @internal List metadata and its single-use source are captured before any item is enumerated. */
export interface PdfCapturedListOptions {
  readonly items: Iterable<PdfListItem> | AsyncIterable<PdfListItem>;
  readonly ordered: boolean;
  readonly start: number;
  readonly itemStyle?: string;
  readonly indent?: number;
  readonly markerGap?: number;
}
/** @internal Pure layout; semantic IDs and page placement belong to the existing page writer. */
export interface PdfPreparedListItem {
  readonly paragraph: PdfPreparedParagraph;
  readonly marker: PdfInlineLine;
  readonly indent: number;
  readonly markerGap: number;
}

export function captureListOptions(block: PdfListBlock, settings: PdfSettings): PdfCapturedListOptions {
  const { items, ordered = false, start: inputStart, itemStyle, indent, markerGap } = block;
  const start = inputStart === undefined ? 1 : inputStart;
  if (typeof ordered !== "boolean") throw new TypeError("PDF list ordered must be boolean.");
  if (!Number.isSafeInteger(start) || start < 1) throw new RangeError("PDF list start must be a positive safe integer.");
  if (!ordered && inputStart !== undefined) throw new TypeError("PDF list start requires ordered:true.");
  for (const [name, value] of [["indent", indent], ["markerGap", markerGap]] as const)
    if (value !== undefined && (!Number.isFinite(value) || value < 0 || value > 14400))
      throw new RangeError("PDF list " + name + " must be from zero through 14,400 points.");
  if (resolveParagraphStyle(itemStyle, {}, settings.paragraphStyles).keepWithNext)
    throw new TypeError("PDF list item styles cannot use keepWithNext.");
  return Object.freeze({ items, ordered, start, ...(itemStyle === undefined ? {} : { itemStyle }),
    ...(indent === undefined ? {} : { indent }), ...(markerGap === undefined ? {} : { markerGap }) });
}

/** Capture the whole rich item before layout can yield; marker and body share fonts, bounds and cancellation. */
export async function prepareListItem(item: PdfListItem, index: number, options: PdfCapturedListOptions,
  settings: PdfSettings, fonts: PdfFontResources): Promise<PdfPreparedListItem> {
  if (!item || typeof item !== "object" || Array.isArray(item) || Object.keys(item).some(key => !["text", "style", "keepTogether"].includes(key)))
    throw new TypeError("PDF list items support text, style and keepTogether.");
  if (!Number.isSafeInteger(index) || index < 0 || !Number.isSafeInteger(options.start + index))
    throw new RangeError("PDF list numbering must remain within positive safe integers.");
  const { text, style = options.itemStyle, keepTogether } = item;
  const captured = captureParagraph({ kind: "paragraph", text, ...(style === undefined ? {} : { style }),
    ...(keepTogether === undefined ? {} : { keepTogether }) }, settings, fonts);
  if (captured.keepWithNext) throw new TypeError("PDF list item styles cannot use keepWithNext.");
  if (!captured.text.text.trim().length) throw new TypeError("PDF list items require nonempty text.");
  const markerText = options.ordered ? String(options.start + index) + "." : "•";
  const capturedMarker = capturePdfText(markerText, captured.style, settings, fonts);
  settings.budget.cell(markerText, markerText.length);
  const width = settings.page.width - settings.margins.left - settings.margins.right;
  const marker = (await wrapPdfText(capturedMarker, fonts, captured.size, width, settings, false))[0]!;
  const markerGap = options.markerGap ?? captured.size / 2, indent = options.indent ?? Math.max(captured.size * 2, marker.width + markerGap);
  if (marker.width + markerGap > indent + .001) throw new RangeError("PDF list indent cannot fit its marker and gap.");
  if (indent >= width) throw new RangeError("PDF list indent leaves no body text area.");
  const paragraph = await prepareParagraph(captured, settings, fonts, width - indent), first = paragraph.lines[0]!;
  const lineHeight = Math.max(first.lineHeight, marker.lineHeight), ascent = Math.max(first.ascent, marker.ascent);
  return { marker, indent, markerGap, paragraph: { ...paragraph,
    lines: [{ ...first, lineHeight, ascent }, ...paragraph.lines.slice(1)],
    height: paragraph.height + lineHeight - first.lineHeight, firstHeight: paragraph.spaceBefore + lineHeight } };
}
