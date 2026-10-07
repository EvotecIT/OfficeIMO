import type { Column } from "../core/index.js";
import type { InvalidCharacterPolicy } from "../xml/index.js";
import { escapeOoxmlAttribute, escapeXml, xmlDeclaration } from "../xml/index.js";
import { cellText, columnName } from "./values.js";
import { spreadsheetNamespace } from "./styles.js";
import type { TableOptions, FooterOptions } from "./types.js";
import { ExportCell } from "../core/presentation.js";
import { Cell } from "./values.js";
import { totalFormula } from "./layout.js";

export interface TableDefinition { readonly id: number; readonly name: string; readonly headers: readonly string[]; readonly keys: readonly string[]; readonly options: Readonly<TableOptions>; }

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
  return Object.freeze({ id, name, headers: Object.freeze(headers), keys: Object.freeze(columns.map(c => c.key ?? c.header)), options: Object.freeze({ ...options, style }) });
}

export function tableXml(table: TableDefinition, rowCount: number, policy: InvalidCharacterPolicy, headerRow = 1, footer?: FooterOptions): string {
  const dataRef = "A" + headerRow + ":" + columnName(table.headers.length) + (rowCount + headerRow), ref = "A" + headerRow + ":" + columnName(table.headers.length) + (rowCount + headerRow + (footer ? 1 : 0)), options = table.options;
  const name = escapeOoxmlAttribute(table.name, policy);
  return xmlDeclaration + '<table xmlns="' + spreadsheetNamespace + '" id="' + table.id + '" name="' + name + '" displayName="' + name + '" ref="' + ref + '"' + (footer ? ' totalsRowCount="1"' : ' totalsRowShown="0"') + '>' +
    '<autoFilter ref="' + dataRef + '"/><tableColumns count="' + table.headers.length + '">' +
    table.headers.map((header, i) => {
      const key = table.keys[i]!, operation = footer?.totals && Object.prototype.hasOwnProperty.call(footer.totals, key) ? footer.totals[key] : undefined, supplied = footer?.values?.[i], value = supplied instanceof Cell || supplied instanceof ExportCell ? supplied.value : supplied;
      const custom = operation && ["average", "min", "max"].includes(operation);
      const attributes = operation ? ' totalsRowFunction="' + (custom ? "custom" : operation === "count" ? "countNums" : operation) + '"' : typeof value === "string" ? ' totalsRowLabel="' + escapeOoxmlAttribute(value, policy) + '"' : "";
      return '<tableColumn id="' + (i + 1) + '" name="' + escapeOoxmlAttribute(header, policy) + '"' + attributes +
        (custom ? '><totalsRowFormula>' + escapeXml(totalFormula(operation!, columnName(i + 1), headerRow, rowCount), policy) + '</totalsRowFormula></tableColumn>' : '/>');
    }).join("") +
    '</tableColumns><tableStyleInfo name="' + options.style + '" showFirstColumn="' + (options.firstColumn ? 1 : 0) +
    '" showLastColumn="' + (options.lastColumn ? 1 : 0) + '" showRowStripes="' + (options.bandedRows !== false ? 1 : 0) +
    '" showColumnStripes="' + (options.bandedColumns ? 1 : 0) + '"/></table>';
}
