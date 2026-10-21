import { core, zip, xml, opc, xlsx, csv, ExportCell, writeXlsx, writeXlsxTo, preflightPdfText as rootPreflightPdfText, OfficeIMOError as RootOfficeIMOError, detectFeatures as rootDetectFeatures } from "@evotecit/officeimo";
import { BlobByteSink, ChunkedTextSink, detectFeatures, NotSupportedError } from "@evotecit/officeimo/core";
import { ZipWriter, Crc32 } from "@evotecit/officeimo/zip";
import { XmlWriter, escapeXml } from "@evotecit/officeimo/xml";
import { OpcPackage, relationshipTypes, partUri, ContentTypes } from "@evotecit/officeimo/opc";
import { Workbook, Worksheet, Cell, FormulaCell, StyleRegistry, NumberFormats, saveBlob } from "@evotecit/officeimo/xlsx";
import { writeCsv, writeCsvTo } from "@evotecit/officeimo/csv";
import { writePdf, writePdfTo, writePdfDocument, writePdfDocumentTo, pdfTable, PdfFont, preflightPdfText } from "@evotecit/officeimo/pdf";
import type { PdfTextPreflightOptions, PdfTextPreflightResult, PdfTextIssueCode, PdfTaggedOptions } from "@evotecit/officeimo/pdf";
import type { PdfTableBands } from "@evotecit/officeimo/pdf";
import type { PdfTableBands as RootPdfTableBands } from "@evotecit/officeimo";
import type { PdfDrawing, PdfDrawingBlock, PdfDrawingElement, PdfDrawingFont } from "@evotecit/officeimo/pdf";
import type { PdfPngImage } from "@evotecit/officeimo/pdf";
import type { PdfPngImage as RootPdfPngImage } from "@evotecit/officeimo";
import type { Column, Rows } from "@evotecit/officeimo/core";
import type { XlsxColumn } from "@evotecit/officeimo/xlsx";
import type { ConditionalFormat, ListValidation, ChartRangeData } from "@evotecit/officeimo/xlsx";
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
import { createGroupedReport as createRootGroupedReport, addGroupedXlsx as addRootGroupedXlsx, pdfGroupedTable as rootGroupedTable } from "@evotecit/officeimo";
import type { GroupedReportOptions, GroupedReportLevel, GroupedReportGroup, GroupedReportValueContext, GroupedXlsxWorksheet, GroupedXlsxSheetOptions, GroupedPdfTableOptions, PdfGroupedTableBlock } from "@evotecit/officeimo";
import { createGroupedReport as createCoreGroupedReport } from "@evotecit/officeimo/core";
import type { GroupedReportSource, GroupedReportGroup as CoreReportGroup, GroupedReportValueContext as CoreValueContext } from "@evotecit/officeimo/core";
import { addGroupedXlsx } from "@evotecit/officeimo/xlsx";
import type { GroupedReportSource as XlsxReportSource, GroupedXlsxWorksheet as XlsxAnalysisSheet } from "@evotecit/officeimo/xlsx";
import { pdfGroupedTable } from "@evotecit/officeimo/pdf";
import type { GroupedReportSource as PdfReportSource, PdfGroupedTableBlock as PdfAnalysisBlock } from "@evotecit/officeimo/pdf";
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
  // @ts-expect-error workbook settings belong in xlsx
  await exportDataTable(host, table, 'xlsx', { workbook: { dateMode: 'utc' } });
  await exportDataTable(host, table, 'xlsx', { xlsx: { sheet: { table: {} } }, columnOptions: { 1: { type: 'number', format: '0.00' } } });
  // @ts-expect-error adapter columns use portable presentation, not private workbook IDs
  await exportDataTable(host, table, 'xlsx', { columnOptions: { 0: { style: 1 } } });
  // @ts-expect-error header style IDs belong to the advanced Workbook API
  await exportDataTable(host, table, 'xlsx', { xlsx: { sheet: { headerStyle: 1 } } });
  // @ts-expect-error portable style patches contain definitions, not workbook component indexes
  await exportDataTable(host, table, 'xlsx', { xlsx: { sheet: { alternatingRowStyle: { font: 0 } } } });
  // @ts-expect-error custom adapter writers return portable values
  await exportDataTable(host, table, 'xlsx', { xlsx: { cellValueWriters: { custom: v => new Cell(v, 0) } } });
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

async function exportWidePdfReports() {
  const rows = [{ service: "WEB-000001", owner: "Operations", latency: 4.9 }];
  const columns = [
    { header: "Service", key: "service" },
    { header: "Owner", key: "owner" },
    { header: "Latency", key: "latency", type: "number" }
  ] as const;
  const wideTable: PdfTableBands = { mode: "bands", minimumColumnWidth: 250, repeatColumns: [0] };
  const rootPolicy: RootPdfTableBands = wideTable;
  void rootPolicy;
  const limits = { maxRows: 100, maxPages: 20, maxTableBands: 4, maxTableWindowRows: 8 };
  await writePdf(rows, { columns, wideTable, limits });
  await writePdfTo(rows, new WritableStream<Uint8Array>(), { columns, wideTable, limits });
  await writePdfDocument([
    { kind: "paragraph", text: "Service inventory" },
    pdfTable(rows, { columns, wideTable })
  ], { limits });
  // @ts-expect-error horizontal pagination has one explicit band policy
  pdfTable(rows, { columns, wideTable: { mode: "shrink", repeatColumns: [0] } });
  // @ts-expect-error repeated columns are zero-based indices, not displayed labels
  pdfTable(rows, { columns, wideTable: { mode: "bands", repeatColumns: ["Service"] } });
}
void exportWidePdfReports;

interface RecordRow { readonly name: string; readonly seen: Date; }
const columns = [{ header: "Name", key: "name" }, { header: "Seen", key: "seen", type: "date", format: "yyyy-mm-dd" }] as const satisfies readonly Column<RecordRow>[];
const records: readonly RecordRow[] = [{ name: "Łódź", seen: new Date() }];
const rows: Rows = [{ name: "Łódź", seen: new Date() }];
const book: Workbook = new Workbook({ signal: new AbortController().signal, onProgress: p => console.log(p.rows) });
const style = book.styles.add({ font: { bold: true }, fill: { color: "ABCDEF" }, border: { bottom: { style: "thin" } }, numberFormat: NumberFormats.Date });
const conditional: readonly ConditionalFormat[] = [
  { type: "cellIs", range: { column: "name" }, operator: "equal", value: "Łódź", style: { font: { bold: true } } },
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
nativeChartData.categories[0] = "changed";
// @ts-expect-error arbitrary styling is outside the native chart contract
const unsupportedChartStyle: import("@evotecit/officeimo/xlsx").WorksheetChart = { ...nativeChart, stacking: true };
void [unsupportedChartKind, invalidChartSeries, unsupportedChartStyle];

async function professionalPdfReport(rows: AsyncIterable<{ service: string; latency: number }>, destination: WritableStream<Uint8Array>) {
  const style: import("@evotecit/officeimo/pdf").PdfParagraphStyle = { fontSize: 12, presentation: { bold: true }, spaceAfter: 6 };
  const section: import("@evotecit/officeimo").PdfSectionBlock = { kind: "section", title: "Service review", headingStyle: "heading" };
  const bookmark: import("@evotecit/officeimo/pdf").PdfBookmark = { title: "Actions", level: 2 };
  const finding: import("@evotecit/officeimo").PdfListItem = { text: [{ text: "Assign an owner. ", bold: true }, { text: "Check the next result." }] };
  const caption: import("@evotecit/officeimo/pdf").PdfCaption = { text: "Table 1. Selected services", style: "caption" };
  const blocks: import("@evotecit/officeimo").PdfDocumentBlock[] = [section,
    { kind: "paragraph", text: "Actions", headingLevel: 2, style: "heading", bookmark },
    { kind: "list", ordered: true, items: [finding], itemStyle: "body" },
    { kind: "section", title: "Inventory", level: 2, orientation: "landscape" },
    { ...pdfTable(rows, { columns: [{ header: "Service", key: "service" }, { header: "Latency", key: "latency", type: "number" }] }), caption },
    { kind: "section", title: "Conclusions" }, { kind: "paragraph", text: "Retain this scope.", style: "body" }
  ];
  const options: import("@evotecit/officeimo/pdf").PdfDocumentOptions = { paragraphStyles: { heading: style, body: {}, caption: { fontSize: 9 } },
    pageHeader(context) { return `${context.pageNumber}: ${context.sectionTitle ?? "Report"} / ${context.sectionNumber}.${context.sectionPageNumber}`; },
    limits: { maxSections: 8, maxBookmarks: 32, maxListItems: 100, maxParagraphStyles: 4 } };
  return writePdfDocumentTo(blocks, destination, options);
}
void professionalPdfReport;

async function editableAnalysisReport(worker: Worker) {
  const validation: ListValidation = { type: "list", column: "status", items: ["Healthy", "Offline"], allowBlank: false,
    showErrorMessage: true, errorStyle: "stop", errorTitle: "Choose status", errorMessage: "Select a listed status." };
  const rootValidation: import("@evotecit/officeimo").ListValidation = validation;
  const analysis = new Workbook({ limits: { maxFormulaCharacters: 10000, maxDataValidations: 8, maxValidationTextCharacters: 1000 },
    cellValueWriters: { doubled: value => new FormulaCell("1+1", Number(value) * 2) } });
  const services = analysis.addWorksheet("Services", { columns: [{ header: "Status", key: "status" }], dataValidation: [rootValidation] });
  await services.addRows([{ status: "Healthy" }, { status: "Offline" }]); await services.close();
  const statusRange: string | undefined = services.dataRange("status");
  const count = new FormulaCell(statusRange ? `COUNTIF(${statusRange},"Healthy")` : "0", statusRange ? 1 : 0);
  const rootFormula: import("@evotecit/officeimo").FormulaCell = count;
  interface SummaryRow { readonly status: string; readonly count: FormulaCell; }
  const formulaColumns: readonly XlsxColumn<SummaryRow>[] = [{ header: "Status", key: "status" }, { header: "Services", key: "count", type: "number" }];
  const summary = analysis.addWorksheet<SummaryRow>("Overview", { columns: formulaColumns, table: {}, footer: { values: ["Total", new FormulaCell("SUM(B2:B3)", 2)] } });
  const boundData: ChartRangeData = { kind: "range", sheet: summary, rowCount: 2, categoryColumn: "status", series: [{ column: "count", color: "2563EB" }] };
  const rootBoundData: import("@evotecit/officeimo").ChartRangeData = boundData;
  summary.addChart({ type: "columnClustered", data: rootBoundData, row: 6, column: 1, width: 640, height: 360 });
  await summary.addRows([{ status: "Healthy", count: rootFormula }, { status: "Offline", count: new FormulaCell("1", 1) }]);
  const interfaceRows: readonly SummaryRow[] = [{ status: "Healthy", count: rootFormula }];
  const interfaceColumns: readonly XlsxColumn[] = [{ header: "Status", key: "status" }, { header: "Services", key: "count", type: "number" }];
  const interfaceSheet = analysis.addWorksheet("Interface summary", { columns: interfaceColumns });
  await interfaceSheet.addRows(interfaceRows);
  await analysis.finish();
  // @ts-expect-error formula caches are supplied as finite numeric results, not text
  new FormulaCell("1+1", "2");
  // @ts-expect-error authored formula definitions are immutable
  count.cachedValue = 2;
  // @ts-expect-error workbook-local formulas do not enter the portable XLSX table contract
  writeXlsx([[count]], { columns: [{ header: "Count" }] });
  // @ts-expect-error CSV values remain portable scalars or ExportCells
  writeCsv([[count]], { columns: [{ header: "Count" }] });
  // @ts-expect-error PDF does not evaluate workbook formulas
  writePdf([[count]], { columns: [{ header: "Count" }] });
  // @ts-expect-error worker row bridges carry portable values; the worker job constructs its FormulaCells
  executeWorkerExport(worker, [[count]], new WritableStream<Uint8Array>());
  // @ts-expect-error a FormulaCell cannot be nested inside a scalar styled Cell
  new Cell(count);
  // @ts-expect-error only literal inline-list validation is supported
  const unsupportedValidation: ListValidation = { type: "decimal", column: 1, items: ["1"] };
  // @ts-expect-error validation choices are strings rather than raw numeric values
  const numericChoices: ListValidation = { type: "list", column: 1, items: [1, 2] };
  // @ts-expect-error chart bindings identify their source Worksheet, not a caller-owned range string
  const externalChartSource: ChartRangeData = { kind: "range", sheet: "Services", rowCount: 2, categoryColumn: 1, series: [{ column: 2 }] };
  void [unsupportedValidation, numericChoices, externalChartSource];
}
void editableAnalysisReport;

// @ts-expect-error text comparison supports equality operators only
const unsupportedTextOrdering: ConditionalFormat = { type: "cellIs", range: { column: "name" }, operator: "lessThan", value: "Łódź", style: {} };

async function groupedAnalysis() {
  const columns = [{ header: "Oddział", key: "site" }, { header: "Opóźnienie", key: "latency", type: "number" }] as const;
  const records = [{ site: "Warsaw", latency: 12 }];
  const grouping = { columns, groupBy: "site", totals: { latency: "average" }, limits: { maxRows: 100 } } as const;
  const options = { labelColumn: { header: "Podsumowanie" },
    subtotalLabel: (group: Extract<import("@evotecit/officeimo").GroupedReportEvent, { kind: "groupEnd" }>) => `${group.label}: ${group.rows}`,
    grandTotalLabel: (summary: import("@evotecit/officeimo").GroupedReportSummary) => String(summary.rows) };
  const source = xlsx.createGroupedReport(records, grouping);
  const rootSource: import("@evotecit/officeimo").GroupedReportSource = source;
  // @ts-expect-error grouped sources belong to createGroupedReport, not structural substitutes
  const fakeSource: import("@evotecit/officeimo").GroupedReportSource = { columns: source.columns, events: source.events };
  void fakeSource;
  await xlsx.writeGroupedXlsx(rootSource, { ...options, collapseGroups: true, limits: { maxOutputBytes: 1000000 } });
  await xlsx.writeGroupedXlsxTo(xlsx.createGroupedReport(records, grouping), new WritableStream<Uint8Array>(), options);
  const pdf = await import("@evotecit/officeimo/pdf");
  await pdf.writeGroupedPdf(pdf.createGroupedReport(records, grouping), { ...options,
    pageNumberLabels: { prefix: "Strona ", separator: " z " }, booleanText: { true: "Tak", false: "Nie" } });
  // @ts-expect-error grouped PDF does not offer tagged leaf-table output
  const taggedGroup: import("@evotecit/officeimo").GroupedPdfOptions = { ...options, tagged: { language: "pl" } };
  void taggedGroup;
  await pdf.writeGroupedPdfTo(pdf.createGroupedReport([], grouping), new WritableStream<Uint8Array>(), { ...options, includeGrandTotal: false });
  const blank = new FormulaCell('IF(COUNT(A1:A2)=0,"",AVERAGE(A1:A2))', null);
  void blank;
  // @ts-expect-error output byte ceilings belong to the format writer, not the grouped row source
  xlsx.createGroupedReport(records, { ...grouping, limits: { maxOutputBytes: 0 } });
  // @ts-expect-error stable grouping fields are declared keys, not numeric display-column indices
  xlsx.createGroupedReport(records, { ...grouping, groupBy: 1 });
  // @ts-expect-error a display label is plain text, not a raw boolean
  await writeCsv([[true]], { columns: [{ header: "Ready" }], booleanText: { true: true, false: "Nie" } });
}
void groupedAnalysis;

interface CertificateRecord {
  readonly name: string;
  readonly site: string;
  readonly issuer: string;
  readonly state: "valid" | "expired";
  readonly days: number | null;
}

async function hierarchicalAnalysis(records: readonly CertificateRecord[]) {
  const site: GroupedReportLevel<CertificateRecord> = { key: "site", value(row, context) {
    const sourceContext: CoreValueContext = context;
    const rootContext: GroupedReportValueContext = sourceContext;
    const index: number = rootContext.rowIndex + rootContext.level;
    const field: string = context.key;
    void [index, field];
    // @ts-expect-error grouping contexts are captured read-only metadata
    context.level = 2;
    return row.site;
  }, label: value => String(value) };
  const setup: GroupedReportOptions<CertificateRecord> = {
    columns: [{ key: "name", header: "Certificate" }, { key: "state", header: "State" }, { key: "days", header: "Days", type: "number" }],
    groupBy: [site, { key: "issuer", value: row => row.issuer }, "state"], totals: { days: "average" },
  };
  const source = (): GroupedReportSource => createCoreGroupedReport(records, setup);
  const rootSource = createRootGroupedReport(records, setup);
  const xlsxSource: XlsxReportSource = rootSource, pdfSource: PdfReportSource = rootSource;
  const groupingKeys: readonly string[] = rootSource.groupBy;
  void [xlsxSource, pdfSource, groupingKeys];
  // @ts-expect-error captured grouping keys cannot be modified
  rootSource.groupBy.push("another");
  for await (const event of source().events) {
    if (event.kind === "complete") continue;
    const path: readonly GroupedReportGroup[] = event.path;
    const group: CoreReportGroup | undefined = path.at(-1);
    if (group) { const level: number = group.level; const label: string = group.label; void [level, label]; }
    // @ts-expect-error every event's ancestor path is immutable
    event.path[0] = { groupIndex: 0, key: "x", label: "x", level: 0 };
  }
  const labels: GroupedXlsxSheetOptions = { labelColumn: { header: "Summary" },
    subtotalLabel: group => `${group.path.map(ancestor => ancestor.label).join(" / ")}: ${group.rows}`,
    grandTotalLabel: summary => `${summary.rows} selected certificates`, sheet: { name: "Certificates" } };
  const book = new Workbook({ compression: "store" });
  const cover = book.addWorksheet("Overview", { columns: [{ header: "Report" }] });
  await cover.addRows([["Selected certificate inventory"]]);
  const analysis: GroupedXlsxWorksheet = await addGroupedXlsx(book, source(), labels);
  const subpathSheet: XlsxAnalysisSheet = analysis;
  const physicalRow: number | undefined = analysis.detailRanges[0]?.firstRow;
  const groupPath: readonly GroupedReportGroup[] | undefined = analysis.groups[0]?.path;
  void [subpathSheet, physicalRow, groupPath, analysis.grandSummary.totals];
  await addRootGroupedXlsx(book, source(), { ...labels, sheet: { name: "Appendix" } });
  await book.toBlob();
  const tableOptions: GroupedPdfTableOptions = { labelColumn: labels.labelColumn, subtotalLabel: labels.subtotalLabel,
    grandTotalLabel: summary => String(summary.rows), continuationLabel: group => group.path.map(ancestor => ancestor.label).join(" / ") };
  const block: PdfGroupedTableBlock = pdfGroupedTable(source(), tableOptions);
  const subpathBlock: PdfAnalysisBlock = rootGroupedTable(source(), tableOptions);
  await writePdfDocument([{ kind: "paragraph", text: "Selected certificate inventory" }, block,
    { kind: "paragraph", text: "Second analysis" }, subpathBlock]);
  // @ts-expect-error hidden grouping getters must return a synchronous scalar
  const asyncLevel: GroupedReportLevel<CertificateRecord> = { key: "site", value: row => Promise.resolve(row.site) };
  // @ts-expect-error hidden grouping getters receive the declared record type
  const wrongRecord: GroupedReportLevel<CertificateRecord> = { key: "site", value: (row: { wrong: string }) => row.wrong };
  // @ts-expect-error grouping labels return text
  const numericLabel: GroupedReportLevel<CertificateRecord> = { key: "site", label: () => 1 };
  // @ts-expect-error document geometry belongs to writePdfDocument, not an analytical table block
  pdfGroupedTable(source(), { ...tableOptions, orientation: "landscape" });
  void [asyncLevel, wrongRecord, numericLabel];
}
void hierarchicalAnalysis;

const pdfTotal: import("@evotecit/officeimo").PdfTotal = "sum";
const rootError = new RootOfficeIMOError("INVALID_STATE", "Stopped");
void [pdfTotal, rootError.code, rootDetectFeatures()];
// @ts-expect-error PDF paragraph alignment is left, center or right
const unsupportedParagraphAlignment: import("@evotecit/officeimo").PdfParagraphStyle = { presentation: { alignment: "justify" } };
void unsupportedParagraphAlignment;
