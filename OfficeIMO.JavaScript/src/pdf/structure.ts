import { OfficeIMOError } from "../core/errors.js";
import { checkAbort } from "../core/iteration.js";
import { unicodeHex } from "./objects.js";
import type { PdfObjects } from "./objects.js";
import type { PdfSettings } from "./settings.js";

interface StructureElement {
  readonly id: number;
  readonly parent: number;
  readonly role: string;
  readonly children: string[];
  readonly attributes: string;
}

/** @internal Logical elements outlive page painting so split paragraphs/cells keep one identity. */
export class PdfStructure {
  readonly root: number;
  readonly document: number;
  private readonly parentTree: number;
  private readonly elements = new Map<number, StructureElement>();
  private readonly pageEntries: string[] = [];
  private pageParents: number[] = [];
  private pageId = 0;
  private pageCount = 0;
  private references = 0;
  constructor(private readonly objects: PdfObjects, private readonly settings: PdfSettings) {
    this.root = objects.reserve(); this.parentTree = objects.reserve();
    this.document = this.element("Document", this.root);
  }
  element(role: string, parent = this.document, attributes = ""): number {
    checkAbort(this.settings.options.signal);
    if (this.elements.size >= this.settings.limits.maxStructureElements)
      throw new OfficeIMOError("RESOURCE_LIMIT", "maxStructureElements exceeded.");
    const id = this.objects.reserve(), element = { id, parent, role, children: [], attributes };
    this.elements.set(id, element);
    this.elements.get(parent)?.children.push(id + " 0 R");
    return id;
  }
  tableRow(table: number, headers: readonly boolean[]): readonly number[] {
    const row = this.element("TR", table);
    return headers.map(header => this.element(header ? "TH" : "TD", row, header ? " /A << /O /Table /Scope /Column >>" : ""));
  }
  beginPage(pageId: number): void { this.pageId = pageId; this.pageParents = []; this.pageCount++; }
  /** Every MCID is local to its page; its parent entry points back to the owning logical element. */
  mark(elementId: number): string {
    if (++this.references > this.settings.limits.maxMarkedContentReferences)
      throw new OfficeIMOError("RESOURCE_LIMIT", "maxMarkedContentReferences exceeded.");
    const element = this.elements.get(elementId)!;
    const mcid = this.pageParents.length;
    this.pageParents.push(elementId);
    element.children.push("<< /Type /MCR /Pg " + this.pageId + " 0 R /MCID " + mcid + " >>");
    return "/" + element.role + " << /MCID " + mcid + " >> BDC\n";
  }
  get pageDictionary(): string { return " /StructParents " + (this.pageCount - 1) + " /Tabs /S"; }
  get catalogDictionary(): string {
    return " /MarkInfo << /Marked true >> /Lang <feff" + unicodeHex(this.settings.options.tagged!.language) + "> /StructTreeRoot " + this.root + " 0 R";
  }
  async finishPage(): Promise<void> {
    const array = this.objects.reserve();
    await this.objects.object(array, "[" + this.pageParents.map(id => id + " 0 R").join(" ") + "]");
    this.pageEntries.push((this.pageCount - 1) + " " + array + " 0 R");
    this.pageParents = [];
  }
  async finish(): Promise<void> {
    for (const element of this.elements.values()) {
      checkAbort(this.settings.options.signal);
      await this.objects.object(element.id, "<< /Type /StructElem /S /" + element.role + " /P " + element.parent +
        " 0 R /K [" + element.children.join(" ") + "]" + element.attributes + " >>");
    }
    await this.objects.object(this.parentTree, "<< /Nums [" + this.pageEntries.join(" ") + "] >>");
    await this.objects.object(this.root, "<< /Type /StructTreeRoot /K [" + this.document + " 0 R] /ParentTree " +
      this.parentTree + " 0 R /ParentTreeNextKey " + this.pageCount + " >>");
  }
}

/** Validate author intent even in untagged output so a later opt-in cannot silently change meaning. */
export function graphicAlternative(value: { readonly alternativeText?: string; readonly decorative?: boolean }, settings: PdfSettings): string | undefined {
  if (value.decorative !== undefined && typeof value.decorative !== "boolean") throw new TypeError("Graphic decorative must be boolean.");
  if (value.alternativeText !== undefined) {
    if (typeof value.alternativeText !== "string" || !value.alternativeText.trim()) throw new TypeError("Graphic alternativeText must be nonempty text.");
    if (value.alternativeText.length > settings.limits.maxCellCharacters) throw new OfficeIMOError("RESOURCE_LIMIT", "maxCellCharacters exceeded by graphic alternativeText.");
    unicodeHex(value.alternativeText);
    if (value.decorative) throw new TypeError("A decorative graphic cannot also have alternativeText.");
  }
  if (settings.options.tagged && value.alternativeText === undefined && value.decorative !== true)
    throw new TypeError("Tagged PDF graphics require alternativeText or decorative:true.");
  if (settings.options.tagged && value.alternativeText !== undefined) settings.budget.cell(value.alternativeText);
  return value.alternativeText;
}
