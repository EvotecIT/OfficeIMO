import { core, zip, xml, opc, xlsx, csv, ExportCell, writeXlsx, writeXlsxTo } from "@evotecit/officeimo";
import { BlobByteSink, ChunkedTextSink, detectFeatures, NotSupportedError } from "@evotecit/officeimo/core";
import { ZipWriter, Crc32 } from "@evotecit/officeimo/zip";
import { XmlWriter, escapeXml } from "@evotecit/officeimo/xml";
import { OpcPackage, relationshipTypes, partUri, ContentTypes } from "@evotecit/officeimo/opc";
import { Workbook, Worksheet, Cell, StyleRegistry, NumberFormats, saveBlob } from "@evotecit/officeimo/xlsx";
import { writeCsv, writeCsvTo } from "@evotecit/officeimo/csv";
import { writePdf, writePdfTo, PdfFont } from "@evotecit/officeimo/pdf";
import type { Column, Rows } from "@evotecit/officeimo/core";
import type { XlsxColumn } from "@evotecit/officeimo/xlsx";
import type { ConditionalFormat } from "@evotecit/officeimo/xlsx";
import { createDataTablesExport, exportDataTable, writeDataTableTo, registerDataTablesButtons } from "@evotecit/officeimo/integrations/datatables";
import type { DataTablesApi, DataTablesHost } from "@evotecit/officeimo/integrations/datatables";
import type { ExportLink } from "@evotecit/officeimo";
const portableLink: ExportLink = { target: "https://example.com/report", tooltip: "Open report" };
const linkedValue = new ExportCell(12.5, { text: "12.50 USD", link: portableLink });
void writePdf([[linkedValue]], { columns: [{ header: "Amount" }], limits: { maxHyperlinks: 1 } });
// @ts-expect-error captured portable links are immutable
linkedValue.link!.target = "https://example.com/different";
// @ts-expect-error a link target is required
new ExportCell("value", { link: { tooltip: "Missing target" } });

async function exportGrid(host: DataTablesHost, table: DataTablesApi) {
  const source = createDataTablesExport(host, table, { exportOptions: { columns: ':visible', modifier: { selected: null } },
    project: value => new ExportCell(value instanceof ExportCell ? value.value : value, { presentation: { bold: true } }) });
  void source.rowCount;
  await exportDataTable(host, table, 'xlsx', { sheet: { table: {} }, columnOptions: { 1: { type: 'number', format: '0.00' } } });
  // @ts-expect-error adapter columns use portable presentation, not private workbook IDs
  await exportDataTable(host, table, 'xlsx', { columnOptions: { 0: { style: 1 } } });
  // @ts-expect-error header style IDs belong to the advanced Workbook API
  await exportDataTable(host, table, 'xlsx', { sheet: { headerStyle: 1 } });
  // @ts-expect-error portable style patches contain definitions, not workbook component indexes
  await exportDataTable(host, table, 'xlsx', { sheet: { alternatingRowStyle: { font: 0 } } });
  // @ts-expect-error custom adapter writers return portable values
  await exportDataTable(host, table, 'xlsx', { workbook: { cellValueWriters: { custom: v => new Cell(v, 0) } } });
  await writeDataTableTo(host, table, 'csv', { write() {} }, { serverSide: 'loaded', csv: { quote: 'all' } });
  registerDataTablesButtons(host, { save: async (blob, filename) => { void [blob, filename]; } });
  await exportDataTable(host, table, 'pdf', { pdf: { orientation: 'landscape', title: 'Report' } });
}
void exportGrid;
async function exportPdf() {
  const columns: readonly Column<{name:string; amount:number}>[] = [{header:'Name',key:'name'},{header:'Amount',value:r=>new ExportCell(r.amount,{text:r.amount.toFixed(2)})}];
  await writePdf([{name:'Report',amount:1}],{columns,footer:{totals:{Amount:'sum'}}});
  await writePdfTo([{name:'Report',amount:1}],new WritableStream<Uint8Array>(),{columns,fonts:{regular:new PdfFont(new Uint8Array())}});
  // @ts-expect-error PDF getters must resolve scalar/portable values
  await writePdf([{nested:{x:1}}],{columns:[{header:'Nested',key:'nested'}]});
}
void exportPdf;

interface RecordRow { readonly name: string; readonly seen: Date; }
const columns = [{ header: "Name", key: "name" }, { header: "Seen", key: "seen", type: "date", format: "yyyy-mm-dd" }] as const satisfies readonly Column<RecordRow>[];
const records: readonly RecordRow[] = [{ name: "Łódź", seen: new Date() }];
const rows: Rows = [{ name: "Łódź", seen: new Date() }];
const book: Workbook = new Workbook({ signal: new AbortController().signal, onProgress: p => console.log(p.rows) });
const style = book.styles.add({ font: { bold: true }, fill: { color: "ABCDEF" }, border: { bottom: { style: "thin" } }, numberFormat: NumberFormats.Date });
const conditional: readonly ConditionalFormat[] = [
  { type: "expression", range: { column: "name", through: "seen" }, formula: '$A3="Łódź"', style: { fill: { color: "C6EFCE" } }, stopIfTrue: true },
  { type: "cellIs", range: { column: "seen" }, operator: "between", values: [45000, 50000], style: { font: { bold: true } } },
  { type: "colorScale", range: { column: "seen" }, stops: [{ threshold: { type: "min" }, color: "F8696B" }, { threshold: { type: "max" }, color: "63BE7B" }] },
  { type: "dataBar", range: { column: "seen" }, color: "638EC6", showValue: false }
];
const sheet: Worksheet = book.addWorksheet("Data", { columns });
const report = book.addWorksheet("Report", { columns, table: { name: "SeenReport", style: "TableStyleMedium9" },
  conditionalFormats: conditional,
  title: { text: "Report", style: { font: { size: 20 }, fill: { color: "D9E1F2" } }, height: 32 },
  headerStyle: style, freezeColumns: 1, rowHeight: 24,
  alternatingRowStyle: { fill: { color: "EAF1F8" } },
  rowStyle: context => context.values[0] === "Łódź" ? { font: { bold: true } } : undefined,
  cellStyle: context => context.columnIndex === 0 ? { verticalAlignment: "center" } : undefined });
await report.addRows(records);
report.addHyperlink({ cell: "A3", target: "https://example.com/report", tooltip: "Open report" });
await sheet.addRows(records); await sheet.addRows(rows); await sheet.addRows([[new Cell("name", style), new Date()]]);
async function* asyncRows() { yield ["row", new Date()] as const; }
await sheet.addRows(asyncRows());
const blob: Blob = await book.toBlob();
const csvBlob: Blob = await writeCsv(records, { columns, delimiter: ";", bom: true });
const sink = new BlobByteSink(); await writeCsvTo(records, sink, { columns });
await writeCsv(records, { columns: [{ header: "Name", key: "name", valueFormatter: (value, context) => context.rowIndex + ": " + value }], quote: "strings", nullValue: "missing" });
const text = new ChunkedTextSink(new BlobByteSink()); await text.write("🧪"); await text.close();
const archive = new ZipWriter(); await archive.add("data.csv", new TextEncoder().encode(await csvBlob.text()));
const writer = new XmlWriter(new BlobByteSink()); await writer.startElement("root"); await writer.text("<&"); await writer.endElement(); await writer.close();
const packageFile = new OpcPackage(); packageFile.addPart({ uri: partUri("/data.xml"), contentType: "application/xml", data: "<data/>" });
packageFile.addRelationship("/", { id: "data", type: relationshipTypes.officeDocument, target: "/data.xml" });
const custom = new Workbook({ cellValueWriters: { milliseconds: value => Number(value) / 1000 } });
await custom.addWorksheet("Custom", { columns: [{ header: "Seconds", type: "milliseconds" }] }).addRows([["1250"]]);
void [core, zip, xml, opc, xlsx, csv, new Crc32(), escapeXml("data"), new ContentTypes(), detectFeatures(), new NotSupportedError("feature"), new StyleRegistry()];
const streamed = new Workbook({ sink: { write(bytes) { void bytes; } }, oversizedText: "preserve", limits: { maxRows: 1000, maxBufferedCharacters: 100000, maxOverflowCharacters: 100000 } });
const resolved = [{ name: new ExportCell(12.5, { text: "12.50 USD", presentation: { background: "E2F0D9", bold: true } }) }];
const streamedSheet = streamed.addWorksheet("Resolved", { columns: [{ header: "Amount", key: "name", groups: ["Metrics"], type: "number" }], autoSize: {},
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
new Workbook({ dateMode: "browser" });
// @ts-expect-error unknown delimiter
writeCsv(records, { columns, delimiter: "|" });
// @ts-expect-error CSV requires a projection
writeCsv(records, {});
// @ts-expect-error plain nested objects are not cell values
sheet.addRows([[{ nested: "value" }]]);
// Unselected nested domain fields are allowed.
await writeCsv([{ name: "DC01", seen: new Date(), nested: { value: 1 } }], { columns });
// @ts-expect-error misspelled object keys cannot widen the inferred row type
writeCsv(records, { columns: [{ header: "Name", key: "naem" }] });
// @ts-expect-error object columns require a key or a value getter
writeXlsx(records, { columns: [{ header: "Name" }] });
// @ts-expect-error XLSX checks literal keys against the actual source type
writeXlsx(records, { columns: [{ header: "Name", key: "naem" }] });
interface DomainRow { readonly person: { readonly name: string }; readonly amount: number; }
const domainRows: readonly DomainRow[] = [{ person: { name: "Łódź" }, amount: 12.5 }];
const advancedColumns: readonly XlsxColumn<DomainRow>[] = [{ header: "Amount", key: "amount", style }];
await book.addWorksheet<DomainRow>("Registered", { columns: advancedColumns, headerStyle: style }).addRows(domainRows);
// @ts-expect-error registered columns are workbook-local, including through a declared variable
writeXlsx(domainRows, { columns: advancedColumns });
// @ts-expect-error CSV uses the portable projection
writeCsv(domainRows, { columns: advancedColumns });
// @ts-expect-error helper patches contain definitions rather than workbook component IDs
writeXlsx(domainRows, { columns: [], sheet: { alternatingRowStyle: { font: 0 } } });
// @ts-expect-error helper writers return portable values
writeXlsx(domainRows, { columns: [], cellValueWriters: { custom: v => new Cell(v, 0) } });
// @ts-expect-error common columns use portable presentation
const privateStyle: Column<DomainRow> = { header: "Amount", key: "amount", style: 1 };
void privateStyle;
// @ts-expect-error one-table helpers do not expose private workbook header IDs
writeXlsx(domainRows, { columns: [], sheet: { headerStyle: 1 } });
interface OptionalRow { text?: string | null; date: Date | null; cell: ExportCell; opaque: unknown; array: number[]; }
const optionalColumns: readonly Column<OptionalRow>[] = [{ header: "Text", key: "text" }, { header: "Date", key: "date" }, { header: "Cell", key: "cell" }];
void optionalColumns;
// @ts-expect-error unknown properties need an explicit scalar projection
const opaqueColumn: Column<OptionalRow> = { header: "Opaque", key: "opaque" };
// @ts-expect-error nested arrays need an explicit scalar projection
const arrayColumn: Column<OptionalRow> = { header: "Array", key: "array" };
void [opaqueColumn, arrayColumn];
// @ts-expect-error literal keys must select portable export values; project nested objects explicitly
writeCsv(domainRows, { columns: [{ header: "Person", key: "person" }] });
// @ts-expect-error XLSX uses the same portable key contract as CSV
writeXlsx(domainRows, { columns: [{ header: "Person", key: "person" }] });
const domainColumns = [{ header: "Name", value: (row: DomainRow) => row.person.name },
  { header: "Amount", key: "amount", format: "0.00" }] as const satisfies readonly Column<DomainRow>[];
await writeXlsx(domainRows, { columns: domainColumns, sheet: { title: { text: "Report" }, table: {} } });
await writeCsv(domainRows, { columns: domainColumns });
interface StyledRow { readonly amount: Cell; }
const styledRows: readonly StyledRow[] = [{ amount: new Cell(123, style) }];
const styledColumns: readonly XlsxColumn<StyledRow>[] = [{ header: "Key", key: "amount" },
  { header: "Getter", value: row => row.amount }];
await book.addWorksheet<StyledRow>("Styled domain", { columns: styledColumns, table: {} }).addRows(styledRows);
// @ts-expect-error portable literal keys cannot select advanced Cells
writeXlsx(styledRows, { columns: [{ header: "Amount", key: "amount" }] });
// @ts-expect-error portable positional columns cannot admit advanced Cells
writeXlsx([[new Cell(123, style)]], { columns: [{ header: "Amount" }] });
// @ts-expect-error CSV positional columns cannot admit advanced Cells
writeCsv([[new Cell(123, style)]], { columns: [{ header: "Amount" }] });
// @ts-expect-error portable getters cannot return advanced Cells
writeXlsx(styledRows, { columns: [{ header: "Amount", value: row => row.amount }] });
const nestedArrays = [[{ amount: 123 }]] as const;
await writeXlsx(nestedArrays, { columns: [{ header: "Amount", value: row => row[0].amount }] });
await writeCsv(nestedArrays, { columns: [{ header: "Amount", value: row => row[0].amount }] });
// @ts-expect-error nested positional values require explicit scalar getters
writeXlsx(nestedArrays, { columns: [{ header: "Amount" }] });
await new Workbook().addWorksheet<readonly [Cell]>("Styled array", { columns: [{ header: "Amount" }] }).addRows([[new Cell(123)]]);
const destination = new WritableStream<Uint8Array>({ write(bytes) { void bytes; } });
const tableResult = await writeXlsxTo(domainRows, destination, { columns: domainColumns });
void [tableResult.rows, tableResult.columns, tableResult.bytes];
await writeCsvTo(domainRows, destination, { columns: domainColumns });
const typedSheet: Worksheet<DomainRow> = new Workbook().addWorksheet<DomainRow>("Domain", { columns: domainColumns });
await typedSheet.addRows(domainRows);
// @ts-expect-error a typed worksheet accepts its declared domain rows
typedSheet.addRows(records);
// @ts-expect-error nested getter results must resolve a scalar or ExportCell
writeCsv(domainRows, { columns: [{ header: "Person", value: row => row.person }] });
// @ts-expect-error getters are synchronous
writeXlsx(domainRows, { columns: [{ header: "Name", value: async row => row.person.name }] });
// @ts-expect-error PDF shares the portable getter contract
writePdf(domainRows, { columns: [{ header: "Name", value: async row => row.person.name }] });
// @ts-expect-error PDF does not accept workbook-local advanced Cells
writePdf([[new Cell(1)]], { columns: [{ header: "Amount" }] });
// @ts-expect-error final names are readonly
sheet.name = "new name";
