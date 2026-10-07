import { core, zip, xml, opc, xlsx, csv, createWorkbook, ExportCell } from "@evotecit/officeimo";
import { BlobByteSink, ChunkedTextSink, detectFeatures, NotSupportedError } from "@evotecit/officeimo/core";
import { ZipWriter, Crc32 } from "@evotecit/officeimo/zip";
import { XmlWriter, escapeXml } from "@evotecit/officeimo/xml";
import { OpcPackage, relationshipTypes, partUri, ContentTypes } from "@evotecit/officeimo/opc";
import { Workbook, Worksheet, Cell, StyleRegistry, NumberFormats, saveBlob } from "@evotecit/officeimo/xlsx";
import { writeCsv, writeCsvTo } from "@evotecit/officeimo/csv";
import type { Column, Rows } from "@evotecit/officeimo/core";

const columns = [{ header: "Name", key: "name" }, { header: "Seen", key: "seen", type: "date", format: "yyyy-mm-dd" }] as const satisfies readonly Column[];
interface RecordRow { readonly name: string; readonly seen: Date; }
const records: readonly RecordRow[] = [{ name: "Łódź", seen: new Date() }];
const rows: Rows = [{ name: "Łódź", seen: new Date() }];
const book: Workbook = createWorkbook({ signal: new AbortController().signal, onProgress: p => console.log(p.rows) });
const style = book.styles.add({ font: { bold: true }, fill: { color: "ABCDEF" }, border: { bottom: { style: "thin" } }, numberFormat: NumberFormats.Date });
const sheet: Worksheet = book.addWorksheet("Data", { columns });
const report = book.addWorksheet("Report", { columns, table: { name: "SeenReport", style: "TableStyleMedium9" },
  title: { text: "Report", style: { font: { size: 20 }, fill: { color: "D9E1F2" } }, height: 32 },
  headerStyle: style, freezeColumns: 1, rowHeight: 24,
  alternatingRowStyle: { fill: { color: "EAF1F8" } },
  rowStyle: context => context.values[0] === "Łódź" ? { font: { bold: true } } : undefined,
  cellStyle: context => context.columnIndex === 1 ? { verticalAlignment: "center" } : undefined });
await report.addRows(records);
report.addHyperlink({ cell: "A3", target: "https://example.com/report", tooltip: "Open report" });
await sheet.addRows(records); await sheet.addRows(rows); await sheet.addRows([[new Cell("name", style), new Date()]]);
async function* asyncRows() { yield ["row", new Date()] as const; }
await sheet.addRows(asyncRows());
const blob: Blob = await book.toBlob();
const csvBlob: Blob = await writeCsv(records, { columns, delimiter: ";", bom: true });
const sink = new BlobByteSink(); await writeCsvTo(records, sink, { columns });
await writeCsv(records, { columns: [{ header: "Name", key: "name", valueFormatter: (value, context) => context.row + ": " + value }], quote: "strings", nullValue: "missing" });
const text = new ChunkedTextSink(new BlobByteSink()); await text.write("🧪"); await text.close();
const archive = new ZipWriter(); await archive.add("data.csv", new TextEncoder().encode(await csvBlob.text()));
const writer = new XmlWriter(new BlobByteSink()); await writer.startElement("root"); await writer.text("<&"); await writer.endElement(); await writer.close();
const packageFile = new OpcPackage(); packageFile.addPart({ uri: partUri("/data.xml"), contentType: "application/xml", data: "<data/>" });
packageFile.addRelationship("/", { id: "data", type: relationshipTypes.officeDocument, target: "/data.xml" });
const custom = new Workbook({ cellValueWriters: { milliseconds: value => Number(value) / 1000 } });
await custom.addSheet("Custom", { columns: [{ header: "Seconds", type: "milliseconds" }] }).addRows([["1250"]]);
void [core, zip, xml, opc, xlsx, csv, new Crc32(), escapeXml("data"), new ContentTypes(), detectFeatures(), new NotSupportedError("feature"), new StyleRegistry()];
const streamed = createWorkbook({ sink: { write(bytes) { void bytes; } }, oversizedText: "preserve", limits: { maxRows: 1000, maxBufferedCharacters: 100000, maxOverflowCharacters: 100000 } });
const resolved = [{ name: new ExportCell(12.5, { text: "12.50 USD", presentation: { background: "E2F0D9", bold: true } }) }];
const streamedSheet = streamed.addSheet("Resolved", { columns: [{ header: "Amount", key: "name", groups: ["Metrics"], type: "number" }], autoSize: {},
  footer: { totals: { name: "sum" } }, print: { repeatHeaders: true, orientation: "landscape", margins: { left: 0.25 } } });
await streamedSheet.addRows(resolved); await streamedSheet.close();
const completion = await streamed.finish(); void completion.bytes;
await writeCsv(resolved, { columns: [{ header: "Amount", key: "name" }], valueMode: "display", limits: { maxOutputBytes: 100000 } });
const streamedZip = new ZipWriter({ write(bytes) { void bytes; } });
const zipPart = await streamedZip.openEntry("data.txt"); await zipPart.write(new TextEncoder().encode("data")); await zipPart.close(); await streamedZip.finish();
const streamedOpc = new OpcPackage({ sink: { write(bytes) { void bytes; } } });
const packagePart = await streamedOpc.openPart("/data.txt", "text/plain"); await packagePart.write(new TextEncoder().encode("data")); await packagePart.close(); await streamedOpc.finish();
function download() { saveBlob(blob, "data.xlsx"); } void download;

// Public input and option guarantees, including domain interfaces without index signatures.
// @ts-expect-error unknown date clock
createWorkbook({ dateMode: "browser" });
// @ts-expect-error unknown delimiter
writeCsv(records, { columns, delimiter: "|" });
// @ts-expect-error CSV requires a projection
writeCsv(records, {});
// @ts-expect-error plain nested objects are not cell values
sheet.addRows([[{ nested: "value" }]]);
// @ts-expect-error unsupported property in a typed object row
writeCsv([{ name: "DC01", nested: { value: 1 } }], { columns });
// @ts-expect-error final names are readonly
sheet.name = "new name";
