import type { Column } from "../core/index.js";
import type { InvalidCharacterPolicy } from "../xml/index.js";
import { escapeOoxmlAttribute, xmlDeclaration } from "../xml/index.js";
import { cellText, columnName } from "./values.js";
import { spreadsheetNamespace } from "./styles.js";
import type { TableOptions } from "./types.js";

export interface TableDefinition { readonly id: number; readonly name: string; readonly headers: readonly string[]; readonly options: Readonly<TableOptions>; }

export function defineTable(id: number, options: TableOptions, columns: readonly Column[], policy: InvalidCharacterPolicy): TableDefinition {
  const name = options.name ?? "Table" + id;
  if (typeof name !== "string" || name.length > 255 || !/^[A-Za-z_][A-Za-z0-9_.]*$/.test(name) ||
      /^(?:[A-Za-z]{1,3}[1-9][0-9]*|R[1-9][0-9]*C[1-9][0-9]*|R|C)$/i.test(name))
    throw new TypeError("Table names must be ASCII identifiers, at most 255 characters, and cannot be cell references.");
  const style = options.style ?? "TableStyleMedium2", match = /^TableStyle(Light|Medium|Dark)([1-9][0-9]?)$/.exec(style);
  if (!match || Number(match[2]) > ({ Light: 21, Medium: 28, Dark: 11 }[match[1]!] ?? 0)) throw new TypeError("Unknown built-in Excel table style.");
  const headers = columns.map(column => cellText(column.header, policy));
  if (!headers.length || headers.some(header => !header.trim() || header.length > 255) || new Set(headers.map(header => header.toLowerCase())).size !== headers.length)
    throw new TypeError("Excel tables require nonblank, unique headers of at most 255 characters.");
  return Object.freeze({ id, name, headers: Object.freeze(headers), options: Object.freeze({ ...options, style }) });
}

export function tableXml(table: TableDefinition, rowCount: number, policy: InvalidCharacterPolicy): string {
  const ref = "A1:" + columnName(table.headers.length) + (rowCount + 1), options = table.options;
  const name = escapeOoxmlAttribute(table.name, policy);
  return xmlDeclaration + '<table xmlns="' + spreadsheetNamespace + '" id="' + table.id + '" name="' + name + '" displayName="' + name + '" ref="' + ref + '" totalsRowShown="0">' +
    '<autoFilter ref="' + ref + '"/><tableColumns count="' + table.headers.length + '">' +
    table.headers.map((header, i) => '<tableColumn id="' + (i + 1) + '" name="' + escapeOoxmlAttribute(header, policy) + '"/>').join("") +
    '</tableColumns><tableStyleInfo name="' + options.style + '" showFirstColumn="' + (options.firstColumn ? 1 : 0) +
    '" showLastColumn="' + (options.lastColumn ? 1 : 0) + '" showRowStripes="' + (options.bandedRows !== false ? 1 : 0) +
    '" showColumnStripes="' + (options.bandedColumns ? 1 : 0) + '"/></table>';
}
