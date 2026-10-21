import { settings } from "./settings.js";
import type { PdfSettings } from "./settings.js";
import type { PdfSectionBlock } from "./section-types.js";

/** @internal Every section derives from captured document defaults, never the preceding section. */
export function sectionSettings(base: PdfSettings, section: Readonly<PdfSectionBlock>): PdfSettings {
  const { pageSize, orientation, margins, pageHeader, pageFooter, pageNumbers } = section;
  const options = { ...base.options,
    ...(pageSize === undefined ? {} : { pageSize }), ...(orientation === undefined ? {} : { orientation }),
    ...(margins === undefined ? {} : { margins }), ...(pageHeader === undefined ? {} : { pageHeader }),
    ...(pageFooter === undefined ? {} : { pageFooter }), ...(pageNumbers === undefined ? {} : { pageNumbers }) };
  // Reuse the existing public normalizer; section changes cannot acquire new limits or reset counters.
  const derived = settings(options, true);
  return { ...base, page: derived.page, margins: derived.margins, options };
}
