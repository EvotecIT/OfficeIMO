import type { PdfObjects } from "./objects.js";

interface ParentNode { readonly id: number; readonly first: number; readonly last: number; }

/** @internal Incremental PDF number tree: at most 64 entries/children per retained level. */
export class PdfParentTree {
  private readonly entries: { readonly key: number; readonly value: number }[] = [];
  private readonly levels: ParentNode[][] = [];
  private lastKey = -1;
  constructor(private readonly objects: PdfObjects, readonly root: number) {}
  async append(key: number, value: number): Promise<void> {
    if (key <= this.lastKey) throw new Error("PDF ParentTree keys must be appended in increasing order.");
    this.lastKey = key; this.entries.push({ key, value });
    if (this.entries.length === 64) await this.leaf();
  }
  private async leaf(): Promise<void> {
    if (!this.entries.length) return;
    const id = this.objects.reserve(), first = this.entries[0]!.key, last = this.entries.at(-1)!.key;
    const body = "<< /Limits [" + first + " " + last + "] /Nums [" + this.entries.map(entry => entry.key + " " + entry.value + " 0 R").join(" ") + "] >>";
    this.entries.length = 0;
    await this.objects.object(id, body); await this.node({ id, first, last }, 0);
  }
  private async node(node: ParentNode, level: number): Promise<void> {
    const nodes = this.levels[level] ??= []; nodes.push(node);
    if (nodes.length === 64) {
      const parent = await this.branch(nodes); nodes.length = 0;
      await this.node(parent, level + 1);
    }
  }
  private async branch(nodes: readonly ParentNode[]): Promise<ParentNode> {
    const id = this.objects.reserve(), first = nodes[0]!.first, last = nodes.at(-1)!.last;
    await this.objects.object(id, "<< /Limits [" + first + " " + last + "] /Kids [" + nodes.map(node => node.id + " 0 R").join(" ") + "] >>");
    return { id, first, last };
  }
  async finish(): Promise<void> {
    await this.leaf();
    // Lower levels contain the newest tail; combine it upward after all older siblings.
    for (let level = 0; level < this.levels.length - 1; level++) {
      const nodes = this.levels[level]!;
      if (nodes.length) { const parent = await this.branch(nodes); nodes.length = 0; await this.node(parent, level + 1); }
    }
    const nodes = this.levels.at(-1) ?? [];
    await this.objects.object(this.root, nodes.length ? "<< /Kids [" + nodes.map(node => node.id + " 0 R").join(" ") + "] >>" : "<< /Nums [] >>");
  }
}
