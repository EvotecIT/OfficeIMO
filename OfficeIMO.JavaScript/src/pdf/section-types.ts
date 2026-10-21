import type { PdfOptions } from "./types.js";
import type { PdfSettings } from "./settings.js";
import { OfficeIMOError } from "../core/errors.js";
import { unicodeHex } from "./objects.js";
import { captureFields } from "../internal/capture.js";

/** A native PDF outline entry whose destination follows the paragraph's actual painted position. */
export interface PdfBookmark {
  readonly title: string;
  readonly level?: 1 | 2 | 3 | 4 | 5 | 6;
}
/** A flat section boundary. Following blocks use this section until the next boundary. */
export interface PdfSectionBlock extends Pick<PdfOptions, "pageSize" | "orientation" | "margins" | "pageHeader" | "pageFooter" | "pageNumbers"> {
  readonly kind: "section";
  readonly title: string;
  readonly level?: 1 | 2 | 3 | 4 | 5 | 6;
  readonly includeHeading?: boolean;
  readonly bookmark?: boolean;
  readonly headingStyle?: string;
}
/** @internal Capture bookmark metadata before source enumeration or cooperative layout. */
export function captureBookmark(value: PdfBookmark | undefined, settings: PdfSettings): Readonly<PdfBookmark> | undefined {
  if (value === undefined) return undefined;
  if (!value || typeof value !== "object") throw new TypeError("PDF bookmark must provide a title and optional level.");
  const { title, level = 1 } = value;
  if (typeof title !== "string" || !title.trim()) throw new TypeError("PDF bookmark title must be nonempty text.");
  if (title.length > settings.limits.maxCellCharacters) throw new OfficeIMOError("RESOURCE_LIMIT", "maxCellCharacters exceeded by bookmark title.");
  if (!Number.isInteger(level) || level < 1 || level > 6) throw new TypeError("PDF bookmark level must be an integer from 1 through 6.");
  unicodeHex(title);
  return Object.freeze({ title, level });
}
/** @internal Sections capture only geometry/decorations; fonts, styles and resource policy remain document-owned. */
export function captureSection(block: PdfSectionBlock, settings: PdfSettings): Readonly<PdfSectionBlock> {
  if (!block || block.kind !== "section") throw new TypeError("PDF section boundary required.");
  if (Object.keys(block).some(key => !["kind", "title", "level", "includeHeading", "bookmark", "headingStyle", "pageSize", "orientation", "margins", "pageHeader", "pageFooter", "pageNumbers"].includes(key)))
    throw new TypeError("PDF sections support title, heading/bookmark controls, page geometry and decorations only.");
  const { title, level, includeHeading = true, bookmark = true, headingStyle, pageSize, orientation, margins, pageHeader, pageFooter, pageNumbers } = block;
  const checked = captureBookmark({ title, ...(level === undefined ? {} : { level }) }, settings)!;
  for (const [name, value] of [["includeHeading", includeHeading], ["bookmark", bookmark], ["pageNumbers", pageNumbers]] as const)
    if (value !== undefined && typeof value !== "boolean") throw new TypeError("PDF section " + name + " must be boolean.");
  if (headingStyle !== undefined && (typeof headingStyle !== "string" || !headingStyle.trim())) throw new TypeError("PDF section headingStyle must be nonempty text.");
  return Object.freeze({ kind: "section", title: checked.title, level: checked.level!, includeHeading, bookmark,
    ...(headingStyle === undefined ? {} : { headingStyle }),
    ...(pageSize === undefined ? {} : { pageSize: typeof pageSize === "object" ? Object.freeze(captureFields(pageSize, ["width", "height"])) : pageSize }),
    ...(orientation === undefined ? {} : { orientation }),
    ...(margins === undefined ? {} : { margins: typeof margins === "object" ? Object.freeze(captureFields(margins, ["top", "right", "bottom", "left"])) : margins }),
    ...(pageHeader === undefined ? {} : { pageHeader }), ...(pageFooter === undefined ? {} : { pageFooter }),
    ...(pageNumbers === undefined ? {} : { pageNumbers }) });
}
