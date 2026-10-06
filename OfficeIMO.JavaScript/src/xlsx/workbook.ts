import { checkAbort } from "../core/iteration.js";
import { OfficeIMOError } from "../core/errors.js";
import { ChunkedTextSink } from "../core/sinks.js";
import { OpcPackage, officeRelationshipsNamespace, relationshipTypes, corePropertiesXml } from "../opc/index.js";
import { escapeOoxmlAttribute, cleanXml, xmlDeclaration } from "../xml/index.js";
import { StyleRegistry, spreadsheetNamespace } from "./styles.js";
import { sheetName } from "./values.js";
import { Worksheet } from "./worksheet.js";
import { defineTable, tableXml } from "./table.js";
import { drawingXml, drawingContentType } from "./attachments.js";
import type { CellValueWriter, WorkbookOptions, SheetOptions, ExtraPart } from "./types.js";

const xlsxMime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const formatType = (name: string) => "application/vnd.openxmlformats-officedocument.spreadsheetml." + name + "+xml";

/** Streaming writer model; append rows through Worksheets, then finalize once. */
export class Workbook {
  readonly styles: StyleRegistry;
  private readonly names = new Set<string>();
  private readonly sheets: Worksheet[] = [];
  private readonly tableNames = new Set<string>();
  private tableCount = 0;
  private readonly writers: Readonly<Record<string, CellValueWriter>>;
  private readonly package: OpcPackage;
  private state = "open";
  private result: Promise<Blob> | undefined;
  /** @internal Immutable settings used by the worksheet owner. */
  readonly settings: Readonly<WorkbookOptions & { dateMode: "local" | "utc"; compression: "auto" | "store"; invalidCharacterPolicy: "strip" | "reject" }>;
  constructor(options: WorkbookOptions = {}) {
    const dateMode = options.dateMode ?? "local", compression = options.compression ?? "auto", policy = options.invalidCharacterPolicy ?? "strip";
    if (dateMode !== "local" && dateMode !== "utc") throw new RangeError("dateMode must be local or utc.");
    if (compression !== "auto" && compression !== "store") throw new RangeError("compression must be auto or store.");
    cleanXml("", policy); corePropertiesXml(options, policy);
    this.styles = new StyleRegistry(policy);
    this.settings = Object.freeze({ ...options, dateMode, compression, invalidCharacterPolicy: policy });
    this.writers = Object.freeze({ ...options.cellValueWriters });
    for (const writer of Object.values(this.writers)) if (typeof writer !== "function") throw new TypeError("Cell value writers must be functions.");
    this.package = new OpcPackage({ compression, invalidCharacterPolicy: policy, ...(options.signal ? { signal: options.signal } : {}) });
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
      this.sheets.map((s, i) => '<sheet name="' + escapeOoxmlAttribute(s.name, this.settings.invalidCharacterPolicy) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join("") + '</sheets></workbook>';
  }
  /** @internal */
  assertOpen(): void {
    checkAbort(this.settings.signal);
    if (this.state !== "open") throw new OfficeIMOError("INVALID_STATE", "Workbook is already finalized.");
  }
  /** @internal */
  writerFor(type: string): CellValueWriter | undefined { return Object.prototype.hasOwnProperty.call(this.writers, type) ? this.writers[type] : undefined; }
  get worksheets(): readonly Worksheet[] { return Object.freeze([...this.sheets]); }
  addWorksheet(name: string, options: SheetOptions = {}): Worksheet {
    this.assertOpen();
    if (this.sheets.length >= 65526) throw new RangeError("Workbook has too many sheets.");
    Worksheet.validate(this, options);
    let tableOptions = options.table;
    if (tableOptions && tableOptions.name === undefined) {
      let suffix = this.tableCount + 1;
      while (this.tableNames.has(("Table" + suffix).toLowerCase())) suffix++;
      tableOptions = { ...tableOptions, name: "Table" + suffix };
    }
    const table = tableOptions ? defineTable(this.tableCount + 1, tableOptions, options.columns ?? [], this.settings.invalidCharacterPolicy) : undefined;
    if (table && this.tableNames.has(table.name.toLowerCase())) throw new TypeError("Duplicate Excel table name: " + table.name);
    const sheet = Worksheet.create(this, sheetName(name, this.names, this.settings.invalidCharacterPolicy), options, table);
    if (table) { this.tableNames.add(table.name.toLowerCase()); this.tableCount++; }
    this.sheets.push(sheet); return sheet;
  }
  /** Existing tabular entry point; returns the same Worksheet model as addWorksheet. */
  addSheet(name: string, options: SheetOptions = {}): Worksheet { return this.addWorksheet(name, options); }
  addPart(part: ExtraPart): void {
    this.assertOpen();
    // Reserve now so collisions and invalid names fail at the call site.
    this.package.addPart(part);
    if (part.relationship) {
      const { source, ...rel } = part.relationship;
      this.package.addRelationship(source ?? "/xl/workbook.xml", { ...rel, target: part.uri });
    }
  }
  toBlob(): Promise<Blob> {
    checkAbort(this.settings.signal);
    if (this.result) return this.result;
    this.assertOpen();
    if (this.sheets.some(s => s.isBusy)) throw new OfficeIMOError("INVALID_STATE", "Await addRows before finalizing the workbook.");
    if (!this.sheets.length) this.addWorksheet("Sheet1");
    this.state = "finalizing";
    this.result = (async () => {
      try {
        for (let i = 0; i < this.sheets.length; i++) {
          const sheet = this.sheets[i]!, uri = "/xl/worksheets/sheet" + (i + 1) + ".xml";
          this.package.addPrepared(uri, formatType("worksheet"), await sheet.finish());
          const table = sheet.tableDefinition;
          if (table) {
            const tableUri = "/xl/tables/table" + table.id + ".xml";
            this.addXmlPart(tableUri, formatType("table"), () => tableXml(table, sheet.rowCount, this.settings.invalidCharacterPolicy));
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
        this.package.addRelationship("/", { id: "workbook", type: relationshipTypes.officeDocument, target: "/xl/workbook.xml" });
        for (let i = 0; i < this.sheets.length; i++) this.package.addRelationship("/xl/workbook.xml", { id: "rId" + (i + 1), type: relationshipTypes.worksheet, target: "/xl/worksheets/sheet" + (i + 1) + ".xml" });
        this.package.addRelationship("/xl/workbook.xml", { id: "styles", type: relationshipTypes.styles, target: "/xl/styles.xml" });
        const blob = await this.package.toBlob(xlsxMime);
        this.settings.onProgress?.({ phase: "complete", rows: this.sheets.reduce((sum, s) => sum + s.rowCount, 0), bytes: blob.size });
        checkAbort(this.settings.signal); this.state = "complete"; return blob;
      } catch (error) {
        this.state = "failed";
        for (const sheet of this.sheets) await sheet.discard(error);
        throw error;
      }
    })();
    return this.result;
  }
}

export function createWorkbook(options: WorkbookOptions = {}): Workbook { return new Workbook(options); }
