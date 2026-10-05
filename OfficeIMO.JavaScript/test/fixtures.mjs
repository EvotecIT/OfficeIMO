import { readFile } from "node:fs/promises";
import { Workbook, Cell } from "../dist/xlsx/index.js";

export const fixtures = JSON.parse(await readFile(new URL("../../OfficeIMO.TestAssets/JavaScript/xlsx-writer.json", import.meta.url), "utf8"));
export async function createFixture(spec, compression) {
  const book = new Workbook({ dateMode: "utc", compression, creator: "OfficeIMO conformance 🧪", title: spec.name,
    created: new Date("2026-10-05T00:00:00Z"), modified: new Date("2026-10-05T00:00:00Z"),
    cellValueWriters: { milliseconds: value => Number(value) / 1000 }, ...spec.options });
  const styles = (spec.styles ?? []).map(s => book.styles.add(s));
  const value = v => v?.kind === "date" ? new Date(v.value) : v?.kind === "cell" ? new Cell(value(v.value), styles[v.style]) : v;
  for (const sheet of spec.sheets) {
    const columns = sheet.columns.map(c => c.style === undefined ? c : { ...c, style: styles[c.style] });
    const worksheet = book.addWorksheet(sheet.name, { ...sheet, columns });
    await worksheet.addRows(sheet.rows.map(row => Array.isArray(row) ? row.map(value) : Object.fromEntries(Object.entries(row).map(([k, v]) => [k, value(v)]))));
  }
  for (const part of spec.parts ?? []) book.addPart(part);
  return book.toBlob();
}
