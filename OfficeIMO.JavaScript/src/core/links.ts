/** A portable external link; the cell's typed value and display text remain independent. */
export interface ExportLink {
  readonly target: string;
  readonly tooltip?: string;
}

/** @internal One target policy shared by document writers and portable cells. */
export function copyExportLink(link: ExportLink): Readonly<ExportLink> {
  if (!link || typeof link !== "object") throw new TypeError("Hyperlinks require an absolute HTTP, HTTPS or mailto target without whitespace or controls.");
  const { target, tooltip } = link;
  if (typeof target !== "string" || !/^(?:https?:\/\/|mailto:)/i.test(target) || /[\u0000-\u0020\u007f]/.test(target))
    throw new TypeError("Hyperlinks require an absolute HTTP, HTTPS or mailto target without whitespace or controls.");
  const url = new URL(target);
  if ((url.protocol === "http:" || url.protocol === "https:") && (!url.hostname || url.username || url.password))
    throw new TypeError("Hyperlink HTTP targets need a host and must not contain credentials.");
  if (tooltip !== undefined && typeof tooltip !== "string") throw new TypeError("Hyperlink tooltip must be text.");
  return Object.freeze({ target: url.href, ...(tooltip === undefined ? {} : { tooltip }) });
}
