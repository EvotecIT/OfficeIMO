import { escapeXml, xmlDeclaration } from "../xml/index.js";
import type { InvalidCharacterPolicy } from "../xml/index.js";

export interface CoreProperties {
  readonly creator?: string;
  readonly title?: string;
  readonly subject?: string;
  readonly description?: string;
  readonly created?: Date;
  readonly modified?: Date;
}
export interface AppProperties { readonly application?: string; readonly company?: string; }

export function corePropertiesXml(properties: CoreProperties = {}, policy: InvalidCharacterPolicy = "strip"): string {
  const created = properties.created ?? new Date(), modified = properties.modified ?? created;
  if (!(created instanceof Date) || !(modified instanceof Date) || !Number.isFinite(created.getTime()) || !Number.isFinite(modified.getTime()))
    throw new TypeError("Invalid workbook property date.");
  return xmlDeclaration + '<cp:coreProperties xmlns:cp="http://schemas.openxmlformats.org/package/2006/metadata/core-properties"' +
    ' xmlns:dc="http://purl.org/dc/elements/1.1/" xmlns:dcterms="http://purl.org/dc/terms/" xmlns:xsi="http://www.w3.org/2001/XMLSchema-instance">' +
    '<dc:creator>' + escapeXml(properties.creator ?? "", policy) + '</dc:creator><dc:title>' + escapeXml(properties.title ?? "", policy) + '</dc:title>' +
    (properties.subject === undefined ? "" : '<dc:subject>' + escapeXml(properties.subject, policy) + '</dc:subject>') +
    (properties.description === undefined ? "" : '<dc:description>' + escapeXml(properties.description, policy) + '</dc:description>') +
    '<dcterms:created xsi:type="dcterms:W3CDTF">' + created.toISOString() + '</dcterms:created>' +
    '<dcterms:modified xsi:type="dcterms:W3CDTF">' + modified.toISOString() + '</dcterms:modified></cp:coreProperties>';
}

export function appPropertiesXml(properties: AppProperties = {}, policy: InvalidCharacterPolicy = "strip"): string {
  return xmlDeclaration + '<Properties xmlns="http://schemas.openxmlformats.org/officeDocument/2006/extended-properties">' +
    '<Application>' + escapeXml(properties.application ?? "OfficeIMO", policy) + '</Application><Company>' + escapeXml(properties.company ?? "", policy) + '</Company></Properties>';
}
