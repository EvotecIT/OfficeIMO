import { createWorkbook, saveBlob } from "@evotecit/officeimo/xlsx";
import { writeCsv } from "@evotecit/officeimo/csv";
import type { Column, Rows } from "@evotecit/officeimo";

const columns = [
  { header: "Name", key: "name", width: 28, type: "string" },
  { header: "Seen", key: "seen", type: "date", format: "yyyy-mm-dd hh:mm" },
  { header: "Healthy", key: "healthy", type: "boolean" }
] as const satisfies readonly Column[];
const rows = [{ name: "DC01", seen: new Date(), healthy: true }];
interface Controller { readonly name: string; readonly seen: Date; readonly healthy: boolean; }
const controllers: readonly Controller[] = rows;
const source: Rows = rows;
const controller = new AbortController();
const book = createWorkbook({ signal: controller.signal, dateMode: "utc", onProgress: p => { console.log(p.rows, p.sheetName, p.bytes); } });
const sheet = book.addSheet("Data", { columns, headerFill: "#D9E1F2", autoFilter: true, freezeHeader: true });
await sheet.addRows(source);
await sheet.addRows(controllers);
async function* asynchronous() { yield ["DC02", new Date(), true] as const; }
await sheet.addRows(asynchronous());
const workbookBlob: Blob = await book.toBlob();
const csvBlob: Blob = await writeCsv(rows, { columns, bom: true, delimiter: ";" });
await writeCsv(controllers, { columns });
function download() { saveBlob(workbookBlob, "data.xlsx"); saveBlob(csvBlob, "data.csv"); }
void download;

// These constraints protect the public input/option/readonly contracts.
// @ts-expect-error unknown date mode
createWorkbook({ dateMode: "browser" });
// @ts-expect-error unknown delimiter
writeCsv(rows, { columns, delimiter: "|" });
// @ts-expect-error a CSV needs declared projection columns
writeCsv(rows, {});
// @ts-expect-error unsupported cell value
sheet.addRows([[{ nested: "value" }]]);
// @ts-expect-error unsupported property in a typed object row
writeCsv([{ name: "DC01", nested: { value: 1 } }], { columns });
// @ts-expect-error final sheet names are readonly
sheet.name = "another";
