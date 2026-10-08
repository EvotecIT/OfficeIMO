import type { ExportLink } from "../core/links.js";
import { OfficeIMOError } from "../core/errors.js";
import { pdfNumber, unicodeHex } from "./objects.js";

const encoder = new TextEncoder();
/** @internal URI actions are byte strings; tooltips are Unicode PDF text strings. */
export function linkAnnotation(link: ExportLink, x: number, bottom: number, width: number, height: number, maxBytes: number): string {
  const prefix = "<< /Type /Annot /Subtype /Link /F 4 /Border [0 0 0] /Rect [" +
    [x, bottom, x + width, bottom + height].map(pdfNumber).join(" ") + "] /A << /S /URI /URI <";
  const suffix = "> >>" + (link.tooltip === undefined ? "" : " /Contents <feff") + " >>";
  const fixedBytes = prefix.length + suffix.length + (link.tooltip === undefined ? 0 : link.tooltip.length * 4 + 1);
  const limit = (): never => { throw new OfficeIMOError("RESOURCE_LIMIT", "maxPageBytes exceeded."); };
  if (fixedBytes + link.target.length * 2 > maxBytes) limit();
  // Count before allocating the UTF-8 bytes or their hex representation. Lone
  // surrogates become U+FFFD, matching TextEncoder's replacement behavior.
  let uriBytes = 0;
  for (let i = 0; i < link.target.length; i++) {
    const code = link.target.charCodeAt(i);
    if (code < 0x80) uriBytes++;
    else if (code < 0x800) uriBytes += 2;
    else if (code >= 0xd800 && code <= 0xdbff && i + 1 < link.target.length &&
      link.target.charCodeAt(i + 1) >= 0xdc00 && link.target.charCodeAt(i + 1) <= 0xdfff) { uriBytes += 4; i++; }
    else uriBytes += 3;
    if (fixedBytes + uriBytes * 2 > maxBytes) limit();
  }
  let uri = "";
  for (const byte of encoder.encode(link.target)) uri += byte.toString(16).padStart(2, "0");
  return prefix + uri + "> >>" +
    (link.tooltip === undefined ? "" : " /Contents <feff" + unicodeHex(link.tooltip) + ">") + " >>";
}
