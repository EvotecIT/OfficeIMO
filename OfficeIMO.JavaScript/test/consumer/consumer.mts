import { core, zip, xml, opc, xlsx, csv, ExportCell, writeXlsx, writeXlsxTo, preflightPdfText as rootPreflightPdfText } from "@evotecit/officeimo";
import { BlobByteSink, ChunkedTextSink, detectFeatures, NotSupportedError } from "@evotecit/officeimo/core";
import { ZipWriter, Crc32 } from "@evotecit/officeimo/zip";
import { XmlWriter, escapeXml } from "@evotecit/officeimo/xml";
import { OpcPackage, relationshipTypes, partUri, ContentTypes } from "@evotecit/officeimo/opc";
import { Workbook, Worksheet, Cell, StyleRegistry, NumberFormats, saveBlob } from "@evotecit/officeimo/xlsx";
import { writeCsv, writeCsvTo } from "@evotecit/officeimo/csv";
import { writePdf, writePdfTo, writePdfDocument, writePdfDocumentTo, pdfTable, PdfFont, preflightPdfText } from "@evotecit/officeimo/pdf";
import type { PdfTextPreflightOptions, PdfTextPreflightResult, PdfTextIssueCode, PdfTaggedOptions } from "@evotecit/officeimo/pdf";
import type { PdfDrawing, PdfDrawingBlock, PdfDrawingElement, PdfDrawingFont } from "@evotecit/officeimo/pdf";
import type { PdfPngImage } from "@evotecit/officeimo/pdf";
import type { PdfPngImage as RootPdfPngImage } from "@evotecit/officeimo";
import type { Column, Rows } from "@evotecit/officeimo/core";
import type { XlsxColumn } from "@evotecit/officeimo/xlsx";
import type { ConditionalFormat } from "@evotecit/officeimo/xlsx";
import { createDataTablesExport, exportDataTable, writeDataTableTo, registerDataTablesButtons } from "@evotecit/officeimo/integrations/datatables";
import type { DataTablesApi, DataTablesHost } from "@evotecit/officeimo/integrations/datatables";
import type { ExportLink } from "@evotecit/officeimo";
import { createCanopyExport, exportCanopy, writeCanopyTo } from "@evotecit/officeimo/integrations/canopyx";
import type { CanopyCapture } from "@evotecit/officeimo/integrations/canopyx";
import { executeWorkerExport } from "@evotecit/officeimo/execution";
import type { WorkerExportOptions } from "@evotecit/officeimo/execution";
import type { PdfFontPreparationOptions } from "@evotecit/officeimo/pdf";
import { createHarfBuzzTextShaper } from "@evotecit/officeimo/pdf/providers/harfbuzz";
import type { PdfHarfBuzzModule, PdfBidiEngine } from "@evotecit/officeimo/pdf/providers/harfbuzz";
import type { PdfTextShaper, PdfShapeLineRequest, PdfShapedRun } from "@evotecit/officeimo";
function optionalTextProvider(harfbuzz: PdfHarfBuzzModule, bidi: PdfBidiEngine, regular: PdfFont, arabic: PdfFont) {
  const provider = createHarfBuzzTextShaper({ harfbuzz, bidi, language: "ar" });
  const replaceable: PdfTextShaper = provider;
  const hostProvider: PdfTextShaper = { shapeLine(request: PdfShapeLineRequest): readonly PdfShapedRun[] { void request; return []; } };
  void [replaceable, hostProvider];
  void writePdfDocument([{ kind: "paragraph", text: [{ text: "Status: ", font: regular }, { text: "مرحبا", font: arabic }] }], { textShaper: provider, textDirection: "auto" });
  // @ts-expect-error text run fonts are explicit registered PdfFont faces, not browser font-family names
  const badFace: import("@evotecit/officeimo").PdfTextRun = { text: "Report", font: "Arial" };
  // @ts-expect-error positioned glyph metrics must be numbers
  const badGlyph: import("@evotecit/officeimo").PdfShapedGlyph = { id: 1, clusterStart: 0, clusterEnd: 1, xAdvance: "wide", yAdvance: 0, xOffset: 0, yOffset: 0 };
  void [badGlyph, badFace];
  provider.dispose();
}
void optionalTextProvider;
async function workerExport(worker: Worker, capture: CanopyCapture, bytes: Uint8Array, signal: AbortSignal) {
  const options: WorkerExportOptions = { signal, batchRows: 64, maxBatchCells: 1024, maxBatchCharacters: 100000, maxChunkBytes: 65536,
    maxSetupBytes: 16000000, maxSetupNodes: 100000 };
  const source = createCanopyExport(capture, "xlsx", { signal });
  await executeWorkerExport(worker, source.rows(), new WritableStream<Uint8Array>(), { ...options, input: { columns: source.columns } });
  const preparation: PdfFontPreparationOptions = { signal, maxBytes: 16000000 };
  const regular = await PdfFont.create(bytes, preparation);
  await writePdfDocument([{ kind: "paragraph", text: "Report" }], { fonts: { regular }, limits: { maxImageWorkingBytes: 64000000 } });
  // @ts-expect-error worker handoff accepts portable projected arrays rather than domain objects
  await executeWorkerExport(worker, [{ value: 1 }], new WritableStream<Uint8Array>());
  // @ts-expect-error font preparation byte limits are numeric
  await PdfFont.create(bytes, { maxBytes: "large" });
}
void workerExport;
async function exportCanopyGrid(capture: CanopyCapture) {
  const source = createCanopyExport(capture, "xlsx", { tones: { warning: { background: "FFF2CC" } } });
  await writeXlsx(source.rows(), { columns: source.columns });
  await exportCanopy(capture, "xlsx", { xlsx: { sheet: { table: {} } }, datetime: "preserve" });
  await writeCanopyTo(capture, "pdf", new WritableStream<Uint8Array>(), { pdf: { orientation: "landscape" } });
  // @ts-expect-error semantic tones use portable presentation rather than private style IDs
  await exportCanopy(capture, "xlsx", { tones: { warning: 1 } });
  // @ts-expect-error capture determines the exported columns
  await exportCanopy(capture, "csv", { csv: { columns: [] } });
  // @ts-expect-error raw/display values belong to the immutable Canopy capture
  await exportCanopy(capture, "csv", { csv: { valueMode: "display" } });
}
void exportCanopyGrid;
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

const preflightOptions: PdfTextPreflightOptions = { bold: true };
const textSupport: PdfTextPreflightResult = preflightPdfText("Report", preflightOptions);
if (!textSupport.supported) {
  const code: PdfTextIssueCode = textSupport.code;
  const sourceIndex: number = textSupport.index;
  const scalar: number | undefined = textSupport.scalar;
  void [code, sourceIndex, scalar, textSupport.message];
}
rootPreflightPdfText("Report");
// @ts-expect-error text preflight selects emphasis with boolean flags
preflightPdfText("Report", { bold: "yes" });

async function exportPdfDocument() {
  const rows = [{ name: "Report", amount: 1 }];
  const table = pdfTable(rows, { columns: [{ header: "Name", key: "name" }, { header: "Amount", value: row => row.amount }], footer: { totals: { Amount: "sum" } } });
  const blocks = [{ kind: "paragraph", text: "Report", fontSize: 18 }, table,
    { kind: "image", image: { format: "rgba", width: 1, height: 1, data: Uint8Array.of(10, 20, 30, 255) } },
    { kind: "vector", width: 20, height: 20, fill: "008000", commands: [{ op: "move", x: 1, y: 1 }, { op: "line", x: 19, y: 19 }] },
    { kind: "pageBreak" }] as const;
  await writePdfDocument(blocks, { title: "Report", limits: { maxImageBytes: 100, maxVectorCommands: 10 } });
  await writePdfDocumentTo(blocks, new WritableStream<Uint8Array>());
  const drawingFont: PdfDrawingFont = "regular";
  const drawingElements: readonly PdfDrawingElement[] = [
    { kind: "path", stroke: "008000", lineCap: "round", lineJoin: "bevel", commands: [{ op: "move", x: -10, y: 0 }, { op: "line", x: 100, y: 50 }] },
    { kind: "group", clip: { x: 0, y: 0, width: 100, height: 50 }, elements: [
      { kind: "text", text: "Diagram", x: 5, y: 20, font: drawingFont, fontSize: 12, color: "111111" }
    ] }
  ];
  const drawing: PdfDrawing = { version: 1, width: 100, height: 50, elements: drawingElements };
  const drawingBlock: PdfDrawingBlock = { kind: "drawing", drawing, alternativeText: "Diagram" };
  await writePdfDocument([drawingBlock], { fonts: { regular: new PdfFont(new Uint8Array()) }, limits: { maxDrawingElements: 100 } });
  // @ts-expect-error portable drawings have one supported wire version
  const unsupportedDrawing: PdfDrawing = { version: 2, width: 100, height: 50, elements: [] };
  void unsupportedDrawing;
  // @ts-expect-error portable text requires an explicit font slot
  const incompleteDrawingText: PdfDrawingElement = { kind: "text", text: "Label", x: 0, y: 0, fontSize: 12, color: "000000" };
  void incompleteDrawingText;
  const png: PdfPngImage = { format: "png", data: new Uint8Array() };
  const rootPng: RootPdfPngImage = png;
  await writePdfDocument([{ kind: "image", image: rootPng }]);
  const tagged: PdfTaggedOptions = { language: "en-US" };
  await writePdfDocument([
    { kind: "paragraph", text: "Sales report", headingLevel: 1, fontSize: 18, presentation: { bold: true } },
    { kind: "paragraph", text: "Services sold this quarter." },
    pdfTable([{ item: "Services", amount: 120 }], { columns: [{ header: "Item", key: "item" }, { header: "Amount", key: "amount" }] }),
    { kind: "image", image: { format: "rgb", width: 1, height: 1, data: Uint8Array.of(0, 128, 200) }, alternativeText: "Blue square" },
    { kind: "vector", width: 10, height: 10, commands: [{ op: "move", x: 0, y: 0 }, { op: "line", x: 10, y: 10 }], decorative: true }
  ], { tagged, limits: { maxStructureElements: 100, maxMarkedContentReferences: 200 } });
  // @ts-expect-error headings use levels from one through six
  await writePdfDocument([{ kind: "paragraph", text: "Heading", headingLevel: 7 }], { tagged });
  // @ts-expect-error table keys are checked against the inferred row type
  pdfTable(rows, { columns: [{ header: "Name", key: "missing" }] });
  // @ts-expect-error table getters must resolve a scalar or portable cell synchronously
  pdfTable(rows, { columns: [{ header: "Name", value: async row => row.name }] });
  // @ts-expect-error raw graphics do not accept arbitrary SVG markup
  await writePdfDocument([{ kind: "vector", width: 20, height: 20, commands: [{ op: "svg", text: "<svg/>" }] }]);
}
void exportPdfDocument;

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

// Native charts accept readonly snapshots through both public entrypoints.
const nativeChartData = {
  categories: ["Łódź", "Warsaw"],
  series: [{ name: "Controllers", values: [12.5, 25], color: "2563EB" }]
} as const satisfies import("@evotecit/officeimo/xlsx").ChartData;
const nativeChart: import("@evotecit/officeimo/xlsx").WorksheetChart = {
  type: "columnClustered", data: nativeChartData, row: 6, column: 1, width: 640, height: 360,
  legend: "bottom", dataLabels: { showValue: true, position: "outsideEnd" },
  categoryAxis: { title: "Site" }, valueAxis: { minimum: 0, maximum: 30, numberFormat: "0.0" }
};
const rootNativeChart: import("@evotecit/officeimo").WorksheetChart = nativeChart;
const chartKinds: readonly import("@evotecit/officeimo").ChartType[] = ["columnClustered", "barClustered", "line", "pie"];
const nativeChartSheet = new Workbook({ limits: { maxCharts: 4, maxChartPoints: 100, maxChartTextCharacters: 10000 } })
  .addWorksheet("Native charts", { columns: [{ header: "Site", key: "site" }, { header: "Actual", key: "actual", type: "number" }], table: {} });
await nativeChartSheet.addRows([{ site: "Łódź", actual: 12.5 }, { site: "Warsaw", actual: 25 }]);
nativeChartSheet.addChart(rootNativeChart);
void chartKinds;
// @ts-expect-error the supported native chart vocabulary is explicit
const unsupportedChartKind: import("@evotecit/officeimo/xlsx").ChartType = "scatter";
// @ts-expect-error native chart values are numeric
const invalidChartSeries: import("@evotecit/officeimo/xlsx").ChartSeries = { name: "Controllers", values: ["12.5"] };
// @ts-expect-error chart input snapshots are readonly
nativeChart.data.categories[0] = "changed";
// @ts-expect-error arbitrary styling is outside the native chart contract
const unsupportedChartStyle: import("@evotecit/officeimo/xlsx").WorksheetChart = { ...nativeChart, stacking: true };
void [unsupportedChartKind, invalidChartSeries, unsupportedChartStyle];
