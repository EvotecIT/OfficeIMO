import { beginTask, checkAbort, pause, taskYieldDue } from "../core/iteration.js";
import { OfficeIMOError } from "../core/errors.js";
import { notifyProgress } from "../core/progress.js";
import { ChunkedTextSink } from "../core/sinks.js";
import { ExportBudget } from "../core/limits.js";
import type { ZipEntry } from "../zip/index.js";
import { OpcPackage, officeRelationshipsNamespace, relationshipTypes, corePropertiesXml } from "../opc/index.js";
import { escapeOoxmlAttribute, cleanXml, xmlDeclaration } from "../xml/index.js";
import type { InvalidCharacterPolicy } from "../xml/index.js";
import { StyleRegistry, spreadsheetNamespace } from "./styles.js";
import { sheetName, assertXlsxValue, copyValue } from "./values.js";
import { copyColumns } from "../internal/rows.js";
import { captureFields } from "../internal/capture.js";
import { captureTableSpans } from "../internal/table-spans.js";
import { assertExportValue } from "../core/presentation.js";
import { Worksheet } from "./worksheet.js";
import { defineTable, tableXml } from "./table.js";
import { copyHyperlink, drawingXml, drawingContentType } from "./attachments.js";
import { TextOverflow } from "./preservation.js";
import type { CellValueWriter, WorkbookOptions, SheetOptions, ExtraPart, XlsxExportResult } from "./types.js";
import { captureWorkbookOptions, sheetOptionFields } from "./options.js";
import { prepareChart } from "./charts.js";
import type { RegisteredChart } from "./charts.js";
import type { WorksheetChart, ChartRangeData } from "./chart-types.js";
import type { ResolvedChartBinding } from "./chart-bindings.js";
import { qualifiedColumnRange, resolveDataColumn } from "./ranges.js";
import { chartContentType, writeChartXml } from "./chart-xml.js";

const xlsxMime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const formatType = (name: string) => "application/vnd.openxmlformats-officedocument.spreadsheetml." + name + "+xml";

/** Capture registered worksheet metadata before preflight and construction inspect it. */
function captureSheet(options: SheetOptions, policy: InvalidCharacterPolicy): SheetOptions {
  const captured = captureFields(options, sheetOptionFields);
  const columns = copyColumns(captured.columns ?? [], true);
  const footer = captured.footer && captureFields(captured.footer, ["rows", "values", "totals", "style"]);
  const print = captured.print && captureFields(captured.print, ["paper", "orientation", "fitToWidth", "fitToHeight", "repeatHeaders", "margins", "header", "footer"]);
  return { ...captured, columns,
    ...(captured.headerRows === undefined ? {} : { headerRows: captureTableSpans(captured.headerRows, columns.length) }),
    ...(footer ? { footer: { ...footer,
      ...(footer.rows === undefined ? {} : { rows: captureTableSpans(footer.rows, columns.length) }),
      ...(footer.values === undefined ? {} : { values: footer.values.map(copyValue) as NonNullable<SheetOptions["footer"]>["values"] & {} }) } } : {}),
    ...(captured.table ? { table: captureFields(captured.table, ["name", "style", "bandedRows", "bandedColumns", "firstColumn", "lastColumn"]) } : {}),
    ...(captured.autoSize ? { autoSize: captureFields(captured.autoSize, ["sampleRows", "minWidth", "maxWidth"]) } : {}),
    ...(captured.title ? { title: captureFields(captured.title, ["text", "style", "height"]) } : {}),
    ...(print ? { print: { ...print, ...(print.margins === undefined ? {} : {
      margins: captureFields(print.margins, ["left", "right", "top", "bottom", "header", "footer"]) }) } } : {}),
    ...(captured.hyperlinks ? { hyperlinks: captured.hyperlinks.map(link => copyHyperlink(link, policy)) } : {}) };
}

/** Streaming writer model; append rows through Worksheets, then finalize once. */
export class Workbook {
  private portableValues = false;
  /** @internal One-table helpers share the writer while admitting only portable values. */
  static forTable(options: WorkbookOptions): Workbook {
    const book = new Workbook(options); book.portableValues = true; return book;
  }
  /** @internal Captured once by the worksheet projector. */
  get valueValidator(): (value: unknown) => void { return this.portableValues ? assertExportValue : assertXlsxValue; }
  readonly styles: StyleRegistry;
  private readonly names = new Set<string>();
  private readonly sheets: Worksheet[] = [];
  private readonly tableNames = new Set<string>();
  private tableCount = 0;
  private readonly writers: Readonly<Record<string, CellValueWriter>>;
  private readonly package: OpcPackage;
  private state = "open";
  private result: Promise<Blob> | undefined;
  private completion: Promise<XlsxExportResult> | undefined;
  private activeSheet: Worksheet | undefined;
  private failure: unknown;
  private overflow: TextOverflow | undefined;
  private links = 0;
  private imageBytes = 0;
  private bufferedBytes = 0;
  private mergedRanges = 0;
  private conditionalFormats = 0;
  private dataValidationCount = 0;
  private validationTextCharacters = 0;
  private formulaCharacters = 0;
  private hasAuthoredFormulas = false;
  private readonly nativeCharts: RegisteredChart[] = [];
  private chartSourceSheets = 0;
  private chartPoints = 0;
  private chartTextCharacters = 0;
  private registering = false;
  /** Caller-owned property getters may run during capture; they cannot register another workbook object. */
  private registration<T>(action: () => T): T {
    this.assertOpen();
    if (this.registering) throw new OfficeIMOError("INVALID_STATE", "Workbook registrations cannot be reentered from input getters.");
    this.registering = true;
    try { return action(); } finally { this.registering = false; }
  }
  /** @internal Shared budget across worksheet headers and data. */
  readonly budget: ExportBudget;
  /** @internal Immutable settings used by the worksheet owner. */
  readonly settings: Readonly<WorkbookOptions & { dateMode: "local" | "utc"; compression: "auto" | "store"; invalidCharacterPolicy: "strip" | "reject" }>;
  constructor(options: WorkbookOptions = {}) {
    options = captureWorkbookOptions(options);
    const dateMode = options.dateMode ?? "local", compression = options.compression ?? "auto", policy = options.invalidCharacterPolicy ?? "strip";
    if (dateMode !== "local" && dateMode !== "utc") throw new RangeError("dateMode must be local or utc.");
    if (compression !== "auto" && compression !== "store") throw new RangeError("compression must be auto or store.");
    if (options.oversizedText !== undefined && !["reject", "preserve"].includes(options.oversizedText)) throw new RangeError("oversizedText must be reject or preserve.");
    cleanXml("", policy); corePropertiesXml(options, policy);
    this.budget = new ExportBudget(options.limits);
    this.styles = new StyleRegistry(policy, options.limits?.maxStyles, options.limits?.maxDifferentialStyles);
    if (options.sink !== undefined && typeof options.sink.write !== "function") throw new TypeError("sink must be a ByteSink.");
    this.settings = Object.freeze({ ...options, ...(options.limits ? { limits: Object.freeze({ ...options.limits }) } : {}), dateMode, compression, invalidCharacterPolicy: policy });
    this.writers = Object.freeze({ ...options.cellValueWriters });
    for (const writer of Object.values(this.writers)) if (typeof writer !== "function") throw new TypeError("Cell value writers must be functions.");
    // One write phase covers every worksheet and incremental append in this book.
    beginTask();
    this.package = new OpcPackage({ compression, invalidCharacterPolicy: policy, ...(options.signal ? { signal: options.signal } : {}),
      ...(options.sink ? { sink: options.sink } : {}), ...(options.limits?.maxOutputBytes === undefined ? {} : { maxOutputBytes: options.limits.maxOutputBytes }) });
    // Property dates and app settings are captured before an asynchronous export begins.
    this.package.setProperties(options, options.appProperties);
    // Register fixed generated parts now; deferred XML sees the final sheets and styles.
    this.addXmlPart("/xl/workbook.xml", formatType("sheet.main"), () => this.workbookXml());
    this.addXmlPart("/xl/styles.xml", formatType("styles"), () => this.styles.toXml());
  }
  private addXmlPart(uri: string, contentType: string, xml: () => string): void {
    this.package.addPart({ uri, contentType, data: async sink => {
      const text = new ChunkedTextSink(sink, this.settings.signal);
      await text.write(xml()); await text.close();
    } });
  }
  private workbookXml(): string {
    return xmlDeclaration + '<workbook xmlns="' + spreadsheetNamespace + '" xmlns:r="' + officeRelationshipsNamespace + '"><workbookPr date1904="0"/><bookViews><workbookView/></bookViews><sheets>' +
      this.sheets.map((s, i) => '<sheet name="' + escapeOoxmlAttribute(s.name, this.settings.invalidCharacterPolicy) + '" sheetId="' + (i + 1) + '"' + (s.hidden ? ' state="hidden"' : "") + ' r:id="rId' + (i + 1) + '"/>').join("") + '</sheets>' +
      (this.sheets.some(s => s.printSettings) ? '<definedNames>' + this.sheets.map((s, i) => {
        const quoted = "'" + s.name.replace(/'/g, "''") + "'!";
        return s.printSettings ? '<definedName name="_xlnm.Print_Area" localSheetId="' + i + '">' + escapeOoxmlAttribute(quoted + '$A$1:$' + s.lastColumn + '$' + s.totalRows, this.settings.invalidCharacterPolicy) + '</definedName>' +
          (s.printSettings.repeatHeaders && s.firstHeaderRow ? '<definedName name="_xlnm.Print_Titles" localSheetId="' + i + '">' + escapeOoxmlAttribute(quoted + '$' + s.firstHeaderRow + ':$' + s.headerRowCount, this.settings.invalidCharacterPolicy) + '</definedName>' : "") : "";
      }).join("") + '</definedNames>' : "") + (this.hasAuthoredFormulas ? '<calcPr calcMode="auto" fullCalcOnLoad="1" forceFullCalc="1"/>' : "") + '</workbook>';
  }
  /** @internal */
  assertOpen(): void {
    try { checkAbort(this.settings.signal); }
    catch (error) { void this.discard(error).catch(() => {}); throw error; }
    if (this.state === "failed") throw this.failure;
    if (this.state !== "open") throw new OfficeIMOError("INVALID_STATE", "Workbook is already finalized.");
  }
  /** @internal */
  writerFor(type: string): CellValueWriter | undefined { return Object.prototype.hasOwnProperty.call(this.writers, type) ? this.writers[type] : undefined; }
  /** @internal */
  preserveText(sheet: string, cell: string, text: string): { preview: string; location: string } {
    if (this.settings.oversizedText !== "preserve") throw new RangeError("Excel cell text exceeds 32,767 UTF-16 code units.");
    if (!this.overflow) {
      this.checkSheetLimit(this.sheets.length + this.chartSourceSheets + 1);
      this.budget.reserve(5, "SheetCellPartTextParts".length);
      this.overflow = new TextOverflow(sheetName("Text overflow", this.names, this.settings.invalidCharacterPolicy), this.settings.limits?.maxOverflowCharacters ?? 4000000, this.settings.invalidCharacterPolicy, (cells, characters, rows) => { this.budget.row(rows); this.budget.reserve(cells, characters); });
    }
    return this.overflow.add(sheet, cell, text);
  }
  private checkSheetLimit(count: number): void {
    if (count > 65526) throw new RangeError("Workbook has too many sheets.");
    if (this.settings.limits?.maxSheets !== undefined && count > this.settings.limits.maxSheets) throw new OfficeIMOError("RESOURCE_LIMIT", "maxSheets exceeded.");
  }
  /** @internal */
  checkLinks(count: number): void {
    if (this.settings.limits?.maxHyperlinks !== undefined && this.links + count > this.settings.limits.maxHyperlinks) throw new OfficeIMOError("RESOURCE_LIMIT", "maxHyperlinks exceeded.");
  }
  /** @internal Reserve a preflighted worksheet batch or one row hyperlink. */
  retainLink(count = 1): void { this.checkLinks(count); this.links += count; }
  /** @internal */
  checkImage(bytes: number): void {
    if (this.settings.limits?.maxImageBytes !== undefined && this.imageBytes + bytes > this.settings.limits.maxImageBytes) throw new OfficeIMOError("RESOURCE_LIMIT", "maxImageBytes exceeded.");
  }
  /** @internal Reserve only a validated image, after checking its budget before copying. */
  retainImage(bytes: number): void {
    this.checkImage(bytes);
    this.imageBytes += bytes;
  }
  /** @internal Bound compressed worksheet retention in Blob mode before final packaging. */
  retainBufferedBytes(bytes: number): void { this.budget.check("maxOutputBytes", this.bufferedBytes + bytes); this.bufferedBytes += bytes; }
  /** @internal Includes generated report merges across all sheets. */
  checkMerges(count: number): void { if (this.mergedRanges + count > (this.settings.limits?.maxMergedRanges ?? 10000)) throw new OfficeIMOError("RESOURCE_LIMIT", "maxMergedRanges exceeded."); }
  /** @internal */
  checkConditionalFormats(count: number): void { if (this.conditionalFormats + count > (this.settings.limits?.maxConditionalFormats ?? 1000)) throw new OfficeIMOError("RESOURCE_LIMIT", "maxConditionalFormats exceeded."); }
  /** @internal Remaining immutable editing-rule budgets, checked before capture allocation. */
  get availableDataValidations(): number { return (this.settings.limits?.maxDataValidations ?? 1000) - this.dataValidationCount; }
  /** @internal */
  get availableValidationTextCharacters(): number { return (this.settings.limits?.maxValidationTextCharacters ?? 1000000) - this.validationTextCharacters; }
  /** @internal Authored formula text is counted once when its cell is serialized. */
  retainFormula(characters: number): void {
    if (this.formulaCharacters + characters > (this.settings.limits?.maxFormulaCharacters ?? 1000000))
      throw new OfficeIMOError("RESOURCE_LIMIT", "maxFormulaCharacters exceeded.");
    this.formulaCharacters += characters; this.hasAuthoredFormulas = true;
  }
  private checkChartText(characters: number): void {
    if (this.chartTextCharacters + characters > (this.settings.limits?.maxChartTextCharacters ?? 1000000))
      throw new OfficeIMOError("RESOURCE_LIMIT", "maxChartTextCharacters exceeded.");
  }
  /** @internal Category caches consume the same cumulative budget as chart metadata. */
  private retainChartText(characters: number): void {
    this.checkChartText(characters);
    this.chartTextCharacters += characters;
  }
  private resolveChartBinding(input: ChartRangeData): { descriptor: ResolvedChartBinding; retainText: (characters: number) => void } {
    const source = input.sheet;
    if (!this.sheets.includes(source) || source.hidden)
      throw new TypeError("Bound charts require a visible worksheet in the same workbook.");
    const metadata = source.chartBindingMetadata, first = input.firstRow ?? 0;
    if (source.isBusy || metadata.closed || first < metadata.rows)
      throw new OfficeIMOError("INVALID_STATE", "Register a chart before its first selected data row is written, while the source worksheet is idle.");
    const category = resolveDataColumn(metadata.columns, input.categoryColumn, "Chart");
    const series = input.series.map(item => resolveDataColumn(metadata.columns, item.column, "Chart"));
    const start = metadata.headerRows + first + 1;
    return { descriptor: { source, sheetName: metadata.name, headerRows: metadata.headerRows, categoryColumn: category, seriesColumns: series,
      seriesNames: series.map(column => metadata.columns[column - 1]!.header),
      categoryFormula: qualifiedColumnRange(metadata.name, category, start, input.rowCount),
      valueFormulas: series.map(column => qualifiedColumnRange(metadata.name, column, start, input.rowCount)),
      nameFormulas: series.map((column, index) => input.series[index]!.name === undefined
        ? metadata.headerReferences[column - 1] : undefined) }, retainText: characters => this.retainChartText(characters) };
  }
  /** @internal Reserve chart snapshots and their source-sheet names before any output is generated. */
  registerChart(input: WorksheetChart): RegisteredChart {
    return this.registration(() => {
      if (this.nativeCharts.length >= (this.settings.limits?.maxCharts ?? 100)) throw new OfficeIMOError("RESOURCE_LIMIT", "maxCharts exceeded.");
      const chart = prepareChart(input, this.settings.invalidCharacterPolicy, (this.settings.limits?.maxChartPoints ?? 100000) - this.chartPoints,
        (this.settings.limits?.maxChartTextCharacters ?? 1000000) - this.chartTextCharacters, this.settings.signal, value => this.resolveChartBinding(value));
      this.assertOpen();
      // Input getters may write categories into an existing bound chart's cache.
      this.checkChartText(chart.textCharacters);
      // Input getters may append rows to an existing sheet and allocate text preservation.
      this.checkSheetLimit(this.sheets.length + this.chartSourceSheets + (chart.binding ? 0 : 1) + (this.overflow ? 1 : 0));
      const registered = Object.freeze({ ...chart, id: this.nativeCharts.length + 1,
        dataSheetName: chart.binding ? (chart.binding.source as Worksheet).name : sheetName("Chart data " + (this.nativeCharts.length + 1), this.names, this.settings.invalidCharacterPolicy) });
      this.nativeCharts.push(registered); this.chartPoints += chart.points; this.chartTextCharacters += chart.textCharacters;
      if (chart.binding) (chart.binding.source as Worksheet).retainChartBinding(chart.binding);
      else this.chartSourceSheets++;
      return registered;
    });
  }
  /** @internal Streamed parts cannot interleave. Starting a new sheet completes the preceding one. */
  async openSheet(sheet: Worksheet): Promise<ZipEntry> {
    if (this.activeSheet && this.activeSheet !== sheet) {
      if (this.activeSheet.isBusy) throw new OfficeIMOError("INVALID_STATE", "Await the active streamed worksheet before starting another.");
      await this.activeSheet.finish();
    }
    this.activeSheet = sheet;
    return this.package.openPart("/xl/worksheets/sheet" + (this.sheets.indexOf(sheet) + 1) + ".xml", formatType("worksheet"));
  }
  /** @internal Preserve the producer's failure while releasing all owned output. */
  async discard(error: unknown): Promise<void> {
    this.state = "failed"; this.failure = error;
    await Promise.allSettled(this.sheets.map(sheet => sheet.discard(error)));
    await this.package.discard(error);
    this.overflow?.clear();
    this.nativeCharts.length = 0;
  }
  get worksheets(): readonly Worksheet[] { return Object.freeze([...this.sheets]); }
  /** Total report data rows; excludes title/header/footer and preservation records. */
  get rowCount(): number { return this.sheets.reduce((sum, sheet) => sum + sheet.reportRowCount, 0); }
  addWorksheet<T = never>(name: string, options: SheetOptions<T> = {}): Worksheet<T> {
    return this.registration(() => {
      this.checkSheetLimit(this.sheets.length + this.chartSourceSheets + 1 + (this.overflow ? 1 : 0));
      // Validation/serialization use erased column metadata; projection later receives T rows.
      const metadata = captureSheet(options as unknown as SheetOptions, this.settings.invalidCharacterPolicy);
      const prepared = Worksheet.validate(this, metadata);
      let tableOptions = metadata.table;
      if (tableOptions && tableOptions.name === undefined) {
        let suffix = this.tableCount + 1;
        while (this.tableNames.has(("Table" + suffix).toLowerCase())) suffix++;
        tableOptions = { ...tableOptions, name: "Table" + suffix };
      }
      const table = tableOptions ? defineTable(this.tableCount + 1, tableOptions, metadata.columns ?? [], this.settings.invalidCharacterPolicy) : undefined;
      if (table && this.tableNames.has(table.name.toLowerCase())) throw new TypeError("Duplicate Excel table name: " + table.name);
      sheetName(name, this.names, this.settings.invalidCharacterPolicy, false);
      const sheet = Worksheet.create(this, () => sheetName(name, this.names, this.settings.invalidCharacterPolicy, false), metadata, table, false, prepared.conditional, prepared.validations);
      this.assertOpen();
      this.checkSheetLimit(this.sheets.length + this.chartSourceSheets + 1 + (this.overflow ? 1 : 0));
      this.names.add(sheet.name.toLowerCase());
      this.mergedRanges += sheet.mergeCount;
      this.conditionalFormats += sheet.conditionalFormatCount;
      this.dataValidationCount += sheet.dataValidations.count;
      this.validationTextCharacters += sheet.dataValidations.textCharacters;
      if (table) { this.tableNames.add(table.name.toLowerCase()); this.tableCount++; }
      this.sheets.push(sheet); return sheet;
    });
  }
  addPart(part: ExtraPart): void {
    this.assertOpen();
    // Reserve now so collisions and invalid names fail at the call site.
    this.package.addPart(part);
    if (part.relationship) {
      const { source, ...rel } = part.relationship;
      this.package.addRelationship(source ?? "/xl/workbook.xml", { ...rel, target: part.uri });
    }
  }
  /** Complete the archive and return counts. Does not close or dispose a caller-owned sink. */
  finish(): Promise<XlsxExportResult> {
    try { checkAbort(this.settings.signal); }
    catch (error) { return this.discard(error).then(() => { throw error; }); }
    if (this.completion) return this.completion;
    if (this.state === "failed") return Promise.reject(this.failure);
    this.assertOpen();
    if (this.sheets.some(s => s.isBusy)) throw new OfficeIMOError("INVALID_STATE", "Await addRows before finalizing the workbook.");
    if (!this.sheets.length) this.addWorksheet("Sheet1");
    this.state = "finalizing";
    this.completion = (async () => {
      try {
        const rows = this.rowCount;
        for (let i = 0; i < this.sheets.length; i++) {
          const sheet = this.sheets[i]!, uri = "/xl/worksheets/sheet" + (i + 1) + ".xml";
          await sheet.finish(); const prepared = sheet.takePrepared();
          if (prepared) this.package.addPrepared(uri, formatType("worksheet"), prepared);
          const table = sheet.tableDefinition;
          if (table) {
            const tableUri = "/xl/tables/table" + table.id + ".xml";
            this.addXmlPart(tableUri, formatType("table"), () => tableXml(table, sheet.rowCount, this.settings.invalidCharacterPolicy, sheet.headerRowCount, sheet.footerSettings));
            this.package.addRelationship(uri, { id: "table", type: officeRelationshipsNamespace + "/table", target: tableUri });
          }
          for (let j = 0; j < sheet.hyperlinks.length; j++) {
            checkAbort(this.settings.signal);
            this.package.addRelationship(uri, { id: "link" + (j + 1), type: officeRelationshipsNamespace + "/hyperlink", target: sheet.hyperlinks[j]!.target, external: true });
            if ((j & 255) === 255 && taskYieldDue()) { await pause(); checkAbort(this.settings.signal); }
          }
          if (sheet.images.length || sheet.charts.length) {
            const drawingUri = "/xl/drawings/drawing" + (i + 1) + ".xml";
            this.addXmlPart(drawingUri, drawingContentType, () => drawingXml(sheet.images, this.settings.invalidCharacterPolicy, sheet.charts));
            this.package.addRelationship(uri, { id: "drawing", type: officeRelationshipsNamespace + "/drawing", target: drawingUri });
            for (let j = 0; j < sheet.images.length; j++) {
              checkAbort(this.settings.signal);
              const imageUri = "/xl/media/sheet" + (i + 1) + "-image" + (j + 1) + ".png";
              this.package.addPart({ uri: imageUri, contentType: "image/png", data: sheet.images[j]!.data });
              this.package.addRelationship(drawingUri, { id: "image" + (j + 1), type: officeRelationshipsNamespace + "/image", target: imageUri });
              if ((j & 255) === 255 && taskYieldDue()) { await pause(); checkAbort(this.settings.signal); }
            }
            for (const chart of sheet.charts) {
              checkAbort(this.settings.signal);
              const chartUri = "/xl/charts/chart" + chart.id + ".xml";
              this.package.addPart({ uri: chartUri, contentType: chartContentType, data: sink => writeChartXml(chart, sink, this.settings.invalidCharacterPolicy, this.settings.signal) });
              this.package.addRelationship(drawingUri, { id: "chart" + chart.id, type: officeRelationshipsNamespace + "/chart", target: chartUri });
            }
          }
        }
        if (this.overflow) {
          const sheet = Worksheet.create(this, this.overflow.sheetName, { columns: [{ header: "Sheet" }, { header: "Cell" }, { header: "Part" }, { header: "Text", width: 80, wrapText: true }, { header: "Parts" }], freezeHeader: true }, undefined, true);
          this.sheets.push(sheet);
          await sheet.finish(this.overflow.values()); const prepared = sheet.takePrepared();
          if (prepared) this.package.addPrepared("/xl/worksheets/sheet" + this.sheets.length + ".xml", formatType("worksheet"), prepared);
          this.overflow.clear();
        }
        for (const chart of this.nativeCharts) {
          checkAbort(this.settings.signal);
          if (chart.binding) continue;
          const sheet = Worksheet.createChartData(this, chart);
          this.sheets.push(sheet);
          function* rows() { for (let i = 0; i < chart.data.categories.length; i++) yield [chart.data.categories[i]!, ...chart.data.series.map(series => series.values[i]!)]; }
          await sheet.finish(rows()); const prepared = sheet.takePrepared();
          if (prepared) this.package.addPrepared("/xl/worksheets/sheet" + this.sheets.length + ".xml", formatType("worksheet"), prepared);
        }
        this.package.addRelationship("/", { id: "workbook", type: relationshipTypes.officeDocument, target: "/xl/workbook.xml" });
        for (let i = 0; i < this.sheets.length; i++) {
          checkAbort(this.settings.signal);
          this.package.addRelationship("/xl/workbook.xml", { id: "rId" + (i + 1), type: relationshipTypes.worksheet, target: "/xl/worksheets/sheet" + (i + 1) + ".xml" });
          if ((i & 255) === 255 && taskYieldDue()) { await pause(); checkAbort(this.settings.signal); }
        }
        this.package.addRelationship("/xl/workbook.xml", { id: "styles", type: relationshipTypes.styles, target: "/xl/styles.xml" });
        const bytes = await this.package.finish();
        for (const sheet of this.sheets) sheet.releaseCharts();
        this.nativeCharts.length = 0;
        notifyProgress(this.settings.onProgress, { phase: "complete", rows, bytes });
        checkAbort(this.settings.signal); this.state = "complete"; return { rows, sheets: this.sheets.length, bytes };
      } catch (error) {
        await this.discard(error);
        throw error;
      }
    })();
    return this.completion;
  }
  toBlob(): Promise<Blob> {
    try { checkAbort(this.settings.signal); }
    catch (error) { return this.discard(error).then(() => { throw error; }); }
    if (this.settings.sink) throw new OfficeIMOError("INVALID_STATE", "This workbook uses a caller-owned sink; use finish().");
    if (!this.result) { const completion = this.finish(); this.result = completion.then(() => this.package.toBlob(xlsxMime)); }
    return this.result;
  }
}
