import { OfficeIMOError } from "../core/errors.js";
import { checkAbort } from "../core/iteration.js";
import { unicodeHex } from "./objects.js";
import type { PdfObjects } from "./objects.js";
import type { PdfSettings } from "./settings.js";
import { PdfParentTree } from "./structure-parent-tree.js";

interface StructureElement {
  readonly id: number;
  readonly parent: number;
  readonly role: string;
  readonly children: (number | string)[];
  readonly linkIdentities: object[];
  readonly attributes: string;
  identifier?: string;
}

interface StructureIdNode { readonly id: number; readonly first: string; readonly last: string; }

/** A sorted name tree retains only a small open leaf and branch at each level. */
class PdfStructureIds {
  private readonly entries: { readonly key: string; readonly value: number }[] = [];
  private readonly levels: StructureIdNode[][] = [];
  private lastKey = "";
  constructor(private readonly objects: PdfObjects, readonly root: number) {}
  async append(key: string, value: number): Promise<void> {
    if (key <= this.lastKey) throw new Error("PDF structure IDs must be appended in increasing order.");
    this.lastKey = key; this.entries.push({ key, value });
    if (this.entries.length === 64) await this.leaf();
  }
  private async leaf(): Promise<void> {
    if (!this.entries.length) return;
    const id = this.objects.reserve(), first = this.entries[0]!.key, last = this.entries.at(-1)!.key;
    const body = "<< /Limits [(" + first + ") (" + last + ")] /Names [" +
      this.entries.map(entry => "(" + entry.key + ") " + entry.value + " 0 R").join(" ") + "] >>";
    this.entries.length = 0;
    await this.objects.object(id, body); await this.node({ id, first, last }, 0);
  }
  private async node(node: StructureIdNode, level: number): Promise<void> {
    const nodes = this.levels[level] ??= []; nodes.push(node);
    if (nodes.length === 64) {
      const parent = await this.branch(nodes); nodes.length = 0; await this.node(parent, level + 1);
    }
  }
  private async branch(nodes: readonly StructureIdNode[]): Promise<StructureIdNode> {
    const id = this.objects.reserve(), first = nodes[0]!.first, last = nodes.at(-1)!.last;
    await this.objects.object(id, "<< /Limits [(" + first + ") (" + last + ")] /Kids [" +
      nodes.map(node => node.id + " 0 R").join(" ") + "] >>");
    return { id, first, last };
  }
  async finish(): Promise<void> {
    await this.leaf();
    for (let level = 0; level < this.levels.length - 1; level++) {
      const nodes = this.levels[level]!;
      if (nodes.length) { const parent = await this.branch(nodes); nodes.length = 0; await this.node(parent, level + 1); }
    }
    const nodes = this.levels.at(-1) ?? [];
    await this.objects.object(this.root, nodes.length ? "<< /Kids [" + nodes.map(node => node.id + " 0 R").join(" ") + "] >>" : "<< /Names [] >>");
  }
}

/** @internal Only open logical elements retain content; completed subtrees retain compact parent references. */
export class PdfStructure {
  readonly root: number;
  readonly document: number;
  private readonly parentTree: number;
  private readonly elements = new Map<number, StructureElement>();
  private readonly parentTreeWriter: PdfParentTree;
  private readonly annotations: { readonly key: number; readonly owner: number }[] = [];
  private readonly links = new WeakMap<object, Map<number, number>>();
  private identifiers: PdfStructureIds | undefined;
  private pageParents: number[] = [];
  private pageId = 0;
  private pageKey = 0;
  private nextParentKey = 0;
  private references = 0;
  private allocated = 0;
  private section: number | undefined;
  constructor(private readonly objects: PdfObjects, private readonly settings: PdfSettings) {
    this.root = objects.reserve(); this.parentTree = objects.reserve();
    this.parentTreeWriter = new PdfParentTree(objects, this.parentTree);
    this.document = this.element("Document", this.root);
  }
  element(role: string, parent = this.section ?? this.document, attributes = ""): number {
    checkAbort(this.settings.options.signal);
    if (++this.allocated > this.settings.limits.maxStructureElements)
      throw new OfficeIMOError("RESOURCE_LIMIT", "maxStructureElements exceeded.");
    const id = this.objects.reserve(), element: StructureElement = { id, parent, role, children: [], attributes, linkIdentities: [] };
    this.elements.set(id, element);
    const owner = this.elements.get(parent);
    // A continuation-only linked caption is discovered during painting, after rows may exist.
    // Its one logical Caption still precedes the Table's rows; other owners retain source order.
    if (role === "Caption" && owner?.role === "Table") owner.children.unshift(id);
    else owner?.children.push(id);
    return id;
  }
  /** Flat visual sections share the same logical Document and cumulative resource policy. */
  async beginSection(): Promise<void> {
    await this.complete(this.section);
    this.section = this.element("Sect", this.document);
  }
  /** Register an explicit header identity in allocation order, without retaining completed headers. */
  async identifier(elementId: number): Promise<string> {
    checkAbort(this.settings.options.signal);
    const element = this.elements.get(elementId);
    if (!element) throw new Error("PDF structure IDs require an open logical owner.");
    if (element.identifier !== undefined) return element.identifier;
    // Object numbers increase monotonically. Fixed-width ASCII preserves name-tree lexical ordering.
    const identifier = "e" + String(elementId).padStart(16, "0");
    this.identifiers ??= new PdfStructureIds(this.objects, this.objects.reserve());
    await this.identifiers.append(identifier, elementId);
    element.identifier = identifier;
    return identifier;
  }
  beginPage(pageId: number): void { this.pageId = pageId; this.pageParents = []; this.pageKey = this.nextParentKey++; }
  /** A source link retains one logical identity when its text/annotation continues on another page. */
  link(parent: number, identity: object, alternativeText?: string): number {
    let parents = this.links.get(identity);
    if (!parents) { parents = new Map(); this.links.set(identity, parents); }
    const existing = parents.get(parent); if (existing !== undefined) return existing;
    const id = this.element("Link", parent, alternativeText === undefined ? "" : " /Alt <feff" + unicodeHex(alternativeText) + ">"); parents.set(parent, id);
    this.elements.get(parent)!.linkIdentities.push(identity); return id;
  }
  /** Associate a native annotation in both directions: OBJR child and annotation ParentTree entry. */
  annotation(elementId: number, annotationId: number): number {
    checkAbort(this.settings.options.signal);
    const element = this.elements.get(elementId);
    if (!element || element.role !== "Link") throw new Error("Tagged link annotations require a logical Link owner.");
    const key = this.nextParentKey++;
    element.children.push("<< /Type /OBJR /Obj " + annotationId + " 0 R /Pg " + this.pageId + " 0 R >>");
    this.annotations.push({ key, owner: elementId });
    return key;
  }
  /** Every MCID is local to its page; its parent entry points back to the owning logical element. */
  mark(elementId: number, actualText?: string): string {
    if (++this.references > this.settings.limits.maxMarkedContentReferences)
      throw new OfficeIMOError("RESOURCE_LIMIT", "maxMarkedContentReferences exceeded.");
    const element = this.elements.get(elementId);
    if (!element) throw new Error("Marked content requires an open logical owner.");
    const mcid = this.pageParents.length;
    this.pageParents.push(elementId);
    element.children.push("<< /Type /MCR /Pg " + this.pageId + " 0 R /MCID " + mcid + " >>");
    return "/" + element.role + " << /MCID " + mcid + (actualText === undefined ? "" : " /ActualText <FEFF" + unicodeHex(actualText) + ">") + " >> BDC\n";
  }
  get pageDictionary(): string { return " /StructParents " + this.pageKey + " /Tabs /S"; }
  get catalogDictionary(): string {
    return " /MarkInfo << /Marked true >> /Lang <feff" + unicodeHex(this.settings.options.tagged!.language) + "> /StructTreeRoot " + this.root + " 0 R";
  }
  async finishPage(): Promise<void> {
    const array = this.objects.reserve();
    await this.objects.object(array, "[" + this.pageParents.map(id => id + " 0 R").join(" ") + "]");
    await this.parentTreeWriter.append(this.pageKey, array);
    for (const annotation of this.annotations) await this.parentTreeWriter.append(annotation.key, annotation.owner);
    this.annotations.length = 0; this.pageParents = [];
  }
  /** Complete only after the last fragment/annotation; split content keeps the same owner until then. */
  async complete(id: number | undefined): Promise<void> {
    if (id === undefined) return;
    const element = this.elements.get(id); if (!element) return;
    for (const child of element.children) if (typeof child === "number" && this.elements.has(child)) await this.complete(child);
    checkAbort(this.settings.options.signal);
    const prefix = "<< /Type /StructElem /S /" + element.role + " /P " + element.parent + " 0 R" +
      (element.identifier === undefined ? "" : " /ID (" + element.identifier + ")") + " /K [";
    const remaining = 32768 - prefix.length - element.attributes.length - 4;
    let children = "", small = remaining >= 0;
    if (small) for (const child of element.children) {
      const token = typeof child === "number" ? child + " 0 R " : child + " ";
      if (children.length + token.length > remaining) { small = false; break; }
      children += token;
    }
    // Most table cells/rows fit in one encoding. Large parents and alternatives remain streamed.
    await this.objects.object(id, small ? prefix + children + "]" + element.attributes + " >>" : this.body(element));
    for (const identity of element.linkIdentities) {
      const parents = this.links.get(identity); parents?.delete(id);
      if (!parents?.size) this.links.delete(identity);
    }
    this.elements.delete(id);
  }
  private *body(element: StructureElement): Iterable<string> {
    yield "<< /Type /StructElem /S /" + element.role + " /P " + element.parent + " 0 R" +
      (element.identifier === undefined ? "" : " /ID (" + element.identifier + ")") + " /K [";
    for (const child of element.children) yield typeof child === "number" ? child + " 0 R " : child + " ";
    yield "]" + element.attributes + " >>";
  }
  async completeRow(cells: readonly number[] | undefined): Promise<void> {
    if (cells?.length) await this.complete(this.elements.get(cells[0]!)?.parent);
  }
  async finish(): Promise<void> {
    await this.complete(this.document);
    await this.parentTreeWriter.finish();
    await this.identifiers?.finish();
    await this.objects.object(this.root, "<< /Type /StructTreeRoot /K [" + this.document + " 0 R] /ParentTree " +
      this.parentTree + " 0 R /ParentTreeNextKey " + this.nextParentKey +
      (this.identifiers ? " /IDTree " + this.identifiers.root + " 0 R" : "") + " >>");
  }
}

/** Validate author intent even in untagged output so a later opt-in cannot silently change meaning. */
export function graphicAlternative(value: { readonly alternativeText?: string; readonly decorative?: boolean }, settings: PdfSettings): string | undefined {
  const decorative = value.decorative, alternativeText = value.alternativeText;
  if (decorative !== undefined && typeof decorative !== "boolean") throw new TypeError("Graphic decorative must be boolean.");
  if (alternativeText !== undefined) {
    if (typeof alternativeText !== "string" || !alternativeText.trim()) throw new TypeError("Graphic alternativeText must be nonempty text.");
    if (alternativeText.length > settings.limits.maxCellCharacters) throw new OfficeIMOError("RESOURCE_LIMIT", "maxCellCharacters exceeded by graphic alternativeText.");
    unicodeHex(alternativeText);
    if (decorative) throw new TypeError("A decorative graphic cannot also have alternativeText.");
  }
  if (settings.options.tagged && alternativeText === undefined && decorative !== true)
    throw new TypeError("Tagged PDF graphics require alternativeText or decorative:true.");
  if (settings.options.tagged && alternativeText !== undefined) settings.budget.cell(alternativeText);
  return alternativeText;
}
