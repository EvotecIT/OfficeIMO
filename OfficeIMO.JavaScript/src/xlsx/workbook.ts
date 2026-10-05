import { checkAbort } from "../core/iteration.js";
import { OfficeIMOError } from "../core/errors.js";
import { OpcPackage, officeRelationshipsNamespace, relationshipTypes, corePropertiesXml } from "../opc/index.js";
import { escapeOoxmlAttribute, cleanXml, xmlDeclaration } from "../xml/index.js";
import { StyleRegistry, spreadsheetNamespace } from "./styles.js";
import { sheetName } from "./values.js";
import { Worksheet } from "./worksheet.js";
import type { CellValueWriter, WorkbookOptions, SheetOptions, ExtraPart } from "./types.js";

const xlsxMime = "application/vnd.openxmlformats-officedocument.spreadsheetml.sheet";
const formatType = (name: string) => "application/vnd.openxmlformats-officedocument.spreadsheetml." + name + "+xml";

/** Streaming writer model; append rows through Worksheets, then finalize once. */
export class Workbook {
  readonly styles: StyleRegistry;
  private readonly names = new Set<string>();
  private readonly sheets: Worksheet[] = [];
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
    const sheet = Worksheet.create(this, sheetName(name, this.names, this.settings.invalidCharacterPolicy), options);
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
        for (let i = 0; i < this.sheets.length; i++) this.package.addPrepared("/xl/worksheets/sheet" + (i + 1) + ".xml", formatType("worksheet"), await this.sheets[i]!.finish());
        this.package.addPart({ uri: "/xl/workbook.xml", contentType: formatType("sheet.main"), data: xmlDeclaration +
          '<workbook xmlns="' + spreadsheetNamespace + '" xmlns:r="' + officeRelationshipsNamespace + '"><workbookPr date1904="0"/><bookViews><workbookView/></bookViews><sheets>' +
          this.sheets.map((s, i) => '<sheet name="' + escapeOoxmlAttribute(s.name, this.settings.invalidCharacterPolicy) + '" sheetId="' + (i + 1) + '" r:id="rId' + (i + 1) + '"/>').join("") + '</sheets></workbook>' });
        this.package.addPart({ uri: "/xl/styles.xml", contentType: formatType("styles"), data: this.styles.toXml() });
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
