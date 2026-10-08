/** A portable external link; the cell's typed value and display text remain independent. */
export interface ExportLink {
  readonly target: string;
  readonly tooltip?: string;
}

/** @internal One target policy shared by document writers and portable cells. */
export function copyExportLink(link: ExportLink): Readonly<ExportLink> {
  if (!link || typeof link !== "object" || typeof link.target !== "string" ||
      !/^(?:https?:\/\/|mailto:)/i.test(link.target) || /[\u0000-\u0020\u007f]/.test(link.target))
    throw new TypeError("Hyperlinks require an absolute HTTP, HTTPS or mailto target without whitespace or controls.");
  const url = new URL(link.target);
  if ((url.protocol === "http:" || url.protocol === "https:") && (!url.hostname || url.username || url.password))
    throw new TypeError("Hyperlink HTTP targets need a host and must not contain credentials.");
  if (link.tooltip !== undefined && typeof link.tooltip !== "string") throw new TypeError("Hyperlink tooltip must be text.");
  return Object.freeze({ target: url.href, ...(link.tooltip === undefined ? {} : { tooltip: link.tooltip }) });
}
