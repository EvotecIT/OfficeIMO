import { readFile } from "node:fs/promises";
import { Workbook, Cell, writeXlsx, writeXlsxTo } from "../dist/xlsx/index.js";
import { ExportCell } from "../dist/core/index.js";

export const fixtures = JSON.parse(await readFile(new URL("../../OfficeIMO.TestAssets/JavaScript/xlsx-writer.json", import.meta.url), "utf8"));
export async function createFixture(spec, compression) {
  if (spec.producer === "table-helper") {
    const sheet = spec.sheets[0], chunks = [];
    const columns = sheet.columns.map(column => column.key === "name" ? { ...column, value: row => row.person.name } : column);
    const rows = sheet.rows.map(row => ({ person: { name: row[0] }, amount: new ExportCell(row[1], { presentation: { background: "C6EFCE" } }), seen: new Date(row[2].value), ignored: { domain: true } }));
    const options = { columns, compression, dateMode: "utc", sheet };
    if (compression === "auto") return writeXlsx(rows, options);
    const stream = new WritableStream({ write: bytes => { chunks.push(new Uint8Array(bytes)); } });
    const result = await writeXlsxTo(rows, stream, options), blob = new Blob(chunks);
    if (stream.locked || result.rows !== rows.length || result.columns !== columns.length || result.bytes !== blob.size) throw new Error("Table helper result/stream ownership differs.");
    return blob;
  }
  const book = new Workbook({ dateMode: "utc", compression, creator: "OfficeIMO conformance 🧪", title: spec.name,
    created: new Date("2026-10-05T00:00:00Z"), modified: new Date("2026-10-05T00:00:00Z"),
    cellValueWriters: { milliseconds: value => Number(value) / 1000 }, ...spec.options });
  const styles = (spec.styles ?? []).map(s => book.styles.add(s));
  const value = v => v?.kind === "date" ? new Date(v.value) : v?.kind === "cell" ? new Cell(value(v.value), styles[v.style]) : v;
  for (const sheet of spec.sheets) {
    const columns = sheet.columns.map(c => c.style === undefined ? c : { ...c, style: styles[c.style] });
    const worksheet = book.addWorksheet(sheet.name, { ...sheet, columns,
      ...(sheet.headerStyle === undefined ? {} : { headerStyle: styles[sheet.headerStyle] }),
      ...(sheet.statusHighlight ? { rowStyle: ({ values }) => values[3] === false ? { fill: { color: "FCE4D6" }, font: { bold: true } } : undefined,
        cellStyle: ({ value, columnIndex }) => columnIndex === 1 && value > 100 ? { font: { color: "C00000" } } : undefined } : {}) });
    await worksheet.addRows(sheet.rows.map(row => Array.isArray(row) ? row.map(value) : Object.fromEntries(Object.entries(row).map(([k, v]) => [k, value(v)]))));
    for (const image of sheet.images ?? []) worksheet.addImage({ ...image, data: Uint8Array.from(atob(image.pngBase64), ch => ch.charCodeAt(0)) });
  }
  for (const part of spec.parts ?? []) book.addPart(part);
  return book.toBlob();
}
