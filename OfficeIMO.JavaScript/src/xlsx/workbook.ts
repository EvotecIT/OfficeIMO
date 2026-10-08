import { checkAbort } from "../core/iteration.js";
import { OfficeIMOError } from "../core/errors.js";
import { ChunkedTextSink } from "../core/sinks.js";
import { ExportBudget } from "../core/limits.js";
import type { ZipEntry } from "../zip/index.js";
import { OpcPackage, officeRelationshipsNamespace, relationshipTypes, corePropertiesXml } from "../opc/index.js";
import { escapeOoxmlAttribute, cleanXml, xmlDeclaration } from "../xml/index.js";
import { StyleRegistry, spreadsheetNamespace } from "./styles.js";
import { sheetName, assertXlsxValue } from "./values.js";
import { assertExportValue } from "../core/presentation.js";
import { Worksheet } from "./worksheet.js";
import { defineTable, tableXml } from "./table.js";
import { drawingXml, drawingContentType } from "./attachments.js";
import { TextOverflow } from "./preservation.js";
import type { CellValueWriter, WorkbookOptions, SheetOptions, ExtraPart, XlsxExportResult } from "./types.js";

const xlsxMime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const formatType = (name: string) => "application/vnd.openxmlformats-officedocument.spreadsheetml." + name + "+xml";

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
  /** @internal Shared budget across worksheet headers and data. */
  readonly budget: ExportBudget;
  /** @internal Immutable settings used by the worksheet owner. */
  readonly settings: Readonly<WorkbookOptions & { dateMode: "local" | "utc"; compression: "auto" | "store"; invalidCharacterPolicy: "strip" | "reject" }>;
  constructor(options: WorkbookOptions = {}) {
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
      this.sheets.map((s, i) => '<sheet name="' + escapeOoxmlAttribute(s.name, this.settings.invalidCharacterPolicy) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join("") + '</sheets>' +
      (this.sheets.some(s => s.printSettings) ? '<definedNames>' + this.sheets.map((s, i) => {
        const quoted = "'" + s.name.replace(/'/g, "''") + "'!";
        return s.printSettings ? '<definedName name="_xlnm.Print_Area" localSheetId="' + i + '">' + escapeOoxmlAttribute(quoted + '$A$1:$' + s.lastColumn + '$' + s.totalRows, this.settings.invalidCharacterPolicy) + '</definedName>' +
          (s.printSettings.repeatHeaders && s.firstHeaderRow ? '<definedName name="_xlnm.Print_Titles" localSheetId="' + i + '">' + escapeOoxmlAttribute(quoted + '$' + s.firstHeaderRow + ':$' + s.headerRowCount, this.settings.invalidCharacterPolicy) + '</definedName>' : "") : "";
      }).join("") + '</definedNames>' : "") + '</workbook>';
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
      this.checkSheetLimit(this.sheets.length + 1);
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
  retainImage(bytes: number): void {
    if (this.settings.limits?.maxImageBytes !== undefined && this.imageBytes + bytes > this.settings.limits.maxImageBytes) throw new OfficeIMOError("RESOURCE_LIMIT", "maxImageBytes exceeded.");
    this.imageBytes += bytes;
  }
  /** @internal Bound compressed worksheet retention in Blob mode before final packaging. */
  retainBufferedBytes(bytes: number): void { this.budget.check("maxOutputBytes", this.bufferedBytes + bytes); this.bufferedBytes += bytes; }
  /** @internal Includes generated report merges across all sheets. */
  checkMerges(count: number): void { if (this.mergedRanges + count > (this.settings.limits?.maxMergedRanges ?? 10000)) throw new OfficeIMOError("RESOURCE_LIMIT", "maxMergedRanges exceeded."); }
  /** @internal */
  checkConditionalFormats(count: number): void { if (this.conditionalFormats + count > (this.settings.limits?.maxConditionalFormats ?? 1000)) throw new OfficeIMOError("RESOURCE_LIMIT", "maxConditionalFormats exceeded."); }
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
  }
  get worksheets(): readonly Worksheet[] { return Object.freeze([...this.sheets]); }
  /** Total report data rows; excludes title/header/footer and preservation records. */
  get rowCount(): number { return this.sheets.reduce((sum, sheet) => sum + sheet.reportRowCount, 0); }
  addWorksheet<T = never>(name: string, options: SheetOptions<T> = {}): Worksheet<T> {
    this.assertOpen();
    this.checkSheetLimit(this.sheets.length + 1 + (this.overflow ? 1 : 0));
    // Validation/serialization use erased column metadata; projection later receives T rows.
    const metadata = options as unknown as SheetOptions;
    const conditional = Worksheet.validate(this, metadata);
    let tableOptions = options.table;
    if (tableOptions && tableOptions.name === undefined) {
      let suffix = this.tableCount + 1;
      while (this.tableNames.has(("Table" + suffix).toLowerCase())) suffix++;
      tableOptions = { ...tableOptions, name: "Table" + suffix };
    }
    const table = tableOptions ? defineTable(this.tableCount + 1, tableOptions, options.columns ?? [], this.settings.invalidCharacterPolicy) : undefined;
    if (table && this.tableNames.has(table.name.toLowerCase())) throw new TypeError("Duplicate Excel table name: " + table.name);
    const sheet = Worksheet.create(this, sheetName(name, this.names, this.settings.invalidCharacterPolicy, false), metadata, table, false, conditional);
    this.names.add(sheet.name.toLowerCase());
    this.mergedRanges += sheet.mergeCount;
    this.conditionalFormats += sheet.conditionalFormatCount;
    if (table) { this.tableNames.add(table.name.toLowerCase()); this.tableCount++; }
    this.sheets.push(sheet); return sheet;
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
        const rows = this.sheets.reduce((sum, s) => sum + s.rowCount, 0);
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
          for (let j = 0; j < sheet.hyperlinks.length; j++) this.package.addRelationship(uri, {
            id: "link" + (j + 1), type: officeRelationshipsNamespace + "/hyperlink", target: sheet.hyperlinks[j]!.target, external: true
          });
          if (sheet.images.length) {
            const drawingUri = "/xl/drawings/drawing" + (i + 1) + ".xml";
            this.addXmlPart(drawingUri, drawingContentType, () => drawingXml(sheet.images, this.settings.invalidCharacterPolicy));
            this.package.addRelationship(uri, { id: "drawing", type: officeRelationshipsNamespace + "/drawing", target: drawingUri });
            for (let j = 0; j < sheet.images.length; j++) {
              const imageUri = "/xl/media/sheet" + (i + 1) + "-image" + (j + 1) + ".png";
              this.package.addPart({ uri: imageUri, contentType: "image/png", data: sheet.images[j]!.data });
              this.package.addRelationship(drawingUri, { id: "image" + (j + 1), type: officeRelationshipsNamespace + "/image", target: imageUri });
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
        this.package.addRelationship("/", { id: "workbook", type: relationshipTypes.officeDocument, target: "/xl/workbook.xml" });
        for (let i = 0; i < this.sheets.length; i++) this.package.addRelationship("/xl/workbook.xml", { id: "rId" + (i + 1), type: relationshipTypes.worksheet, target: "/xl/worksheets/sheet" + (i + 1) + ".xml" });
        this.package.addRelationship("/xl/workbook.xml", { id: "styles", type: relationshipTypes.styles, target: "/xl/styles.xml" });
        const bytes = await this.package.finish();
        this.settings.onProgress?.({ phase: "complete", rows, bytes });
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
