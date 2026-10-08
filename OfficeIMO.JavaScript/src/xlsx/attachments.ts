import { cleanXml, escapeXml, escapeOoxmlAttribute, xmlDeclaration } from "../xml/index.js";
import type { InvalidCharacterPolicy } from "../xml/index.js";
import { copyExportLink } from "../core/links.js";
import { officeRelationshipsNamespace } from "../opc/index.js";

/** A link attached to an existing cell. Its value and presentation remain ordinary cell data. */
export interface Hyperlink { readonly cell: string; readonly target: string; readonly tooltip?: string; }
/** PNG bytes anchored at a one-based row/column, with display size in CSS pixels at 96 DPI. */
export interface WorksheetImage {
  readonly data: Uint8Array;
  readonly row: number;
  readonly column: number;
  readonly width: number;
  readonly height: number;
  readonly description?: string;
}

export function cellPosition(reference: string): { row: number; column: number } {
  const match = typeof reference === "string" && /^([A-Z]{1,3})([1-9][0-9]{0,6})$/.exec(reference);
  if (!match) throw new TypeError("Use a single uppercase A1 cell reference.");
  const row = Number(match[2]); let column = 0;
  for (const letter of match[1]!) column = column * 26 + letter.charCodeAt(0) - 64;
  if (column > 16384 || row > 1048576) throw new RangeError("Cell reference exceeds Excel's worksheet bounds.");
  return { row, column };
}

export function copyHyperlink(link: Hyperlink, policy: InvalidCharacterPolicy): Hyperlink {
  cellPosition(link.cell);
  const copied = copyExportLink(link);
  return Object.freeze({ cell: link.cell, target: copied.target,
    ...(copied.tooltip === undefined ? {} : { tooltip: cleanXml(copied.tooltip, policy) }) });
}

export function copyImage(image: WorksheetImage, policy: InvalidCharacterPolicy): WorksheetImage {
  if (!(image.data instanceof Uint8Array) || image.data.length < 33 ||
      ![137, 80, 78, 71, 13, 10, 26, 10, 0, 0, 0, 13, 73, 72, 68, 82].every((byte, i) => image.data[i] === byte))
    throw new TypeError("Images require PNG bytes with a valid signature and IHDR header.");
  const header = new DataView(image.data.buffer, image.data.byteOffset, image.data.byteLength);
  if (!header.getUint32(16) || !header.getUint32(20)) throw new RangeError("PNG dimensions must be positive.");
  if (!Number.isInteger(image.row) || image.row < 1 || image.row > 1048576 || !Number.isInteger(image.column) || image.column < 1 || image.column > 16384)
    throw new RangeError("Image anchors must be one-based worksheet rows and columns.");
  for (const size of [image.width, image.height]) if (!Number.isFinite(size) || size <= 0 || size > 100000) throw new RangeError("Image display dimensions must be positive and at most 100,000 pixels.");
  return Object.freeze({ ...image, data: new Uint8Array(image.data),
    ...(image.description === undefined ? {} : { description: cleanXml(image.description, policy) }) });
}

export function hyperlinksXml(links: readonly Hyperlink[], policy: InvalidCharacterPolicy, internal: readonly { cell: string; location: string }[] = []): string {
  return '<hyperlinks>' + links.map((link, i) => '<hyperlink ref="' + link.cell + '" r:id="link' + (i + 1) + '"' +
    (link.tooltip === undefined ? "" : ' tooltip="' + escapeOoxmlAttribute(link.tooltip, policy) + '"') + '/>').join("") +
    internal.map(link => '<hyperlink ref="' + link.cell + '" location="' + escapeOoxmlAttribute(link.location, policy) + '"/>').join("") + '</hyperlinks>';
}

export function drawingXml(images: readonly WorksheetImage[], policy: InvalidCharacterPolicy): string {
  const ns = "http://schemas.openxmlformats.org/drawingml/2006/";
  return xmlDeclaration + '<xdr:wsDr xmlns:xdr="' + ns + 'spreadsheetDrawing" xmlns:a="' + ns + 'main" xmlns:r="' + officeRelationshipsNamespace + '">' +
    images.map((image, i) => {
      const size = ' cx="' + Math.round(image.width * 9525) + '" cy="' + Math.round(image.height * 9525) + '"';
      return '<xdr:oneCellAnchor><xdr:from><xdr:col>' + (image.column - 1) + '</xdr:col><xdr:colOff>0</xdr:colOff><xdr:row>' + (image.row - 1) + '</xdr:row><xdr:rowOff>0</xdr:rowOff></xdr:from>' +
        '<xdr:ext' + size + '/><xdr:pic><xdr:nvPicPr><xdr:cNvPr id="' + (i + 1) + '" name="Image ' + (i + 1) + '" descr="' + escapeXml(image.description ?? "", policy) + '"/><xdr:cNvPicPr/></xdr:nvPicPr>' +
        '<xdr:blipFill><a:blip r:embed="image' + (i + 1) + '"/><a:stretch><a:fillRect/></a:stretch></xdr:blipFill>' +
        '<xdr:spPr><a:xfrm><a:off x="0" y="0"/><a:ext' + size + '/></a:xfrm><a:prstGeom prst="rect"><a:avLst/></a:prstGeom></xdr:spPr></xdr:pic><xdr:clientData/></xdr:oneCellAnchor>';
    }).join("") + '</xdr:wsDr>';
}

export const drawingContentType = "application/vnd.openxmlformats-officedocument.drawing+xml";
