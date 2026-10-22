import { OfficeIMOError } from "../core/errors.js";
import { checkAbort } from "../core/iteration.js";
import { pdfNumber, unicodeHex } from "./objects.js";
import type { PdfObjects } from "./objects.js";
import type { PdfSettings } from "./settings.js";
import type { PdfBookmark } from "./section-types.js";

interface OutlineItem {
  readonly id: number; readonly title: string; readonly page: number; readonly y: number;
  readonly level: number; readonly parent: number; readonly children: OutlineItem[];
  readonly previous: number | undefined; next: number | undefined;
}
/** @internal Bounded native bookmarks; destinations are captured only when their content is painted. */
export class PdfOutlines {
  private root: number | undefined;
  private readonly items: OutlineItem[] = [];
  private readonly top: OutlineItem[] = [];
  private readonly stack: OutlineItem[] = [];
  constructor(private readonly objects: PdfObjects, private readonly settings: PdfSettings) {}
  add(bookmark: Readonly<PdfBookmark>, page: number, y: number): void {
    checkAbort(this.settings.options.signal);
    if (this.items.length >= this.settings.limits.maxBookmarks) throw new OfficeIMOError("RESOURCE_LIMIT", "maxBookmarks exceeded.");
    this.root ??= this.objects.reserve();
    const level = bookmark.level ?? 1;
    while (this.stack.length && this.stack.at(-1)!.level >= level) this.stack.pop();
    const parent = this.stack.at(-1), siblings = parent?.children ?? this.top, previous = siblings.at(-1);
    const item: OutlineItem = { id: this.objects.reserve(), title: bookmark.title, page, y, level,
      parent: parent?.id ?? this.root, children: [], previous: previous?.id, next: undefined };
    if (previous) previous.next = item.id;
    siblings.push(item); this.items.push(item); this.stack.push(item);
  }
  get count(): number { return this.items.length; }
  /** A completed analytical group cannot parent later groups, including unnamed ones. */
  closeGroup(level: number): void {
    while (this.stack.length && this.stack.at(-1)!.level >= level) this.stack.pop();
  }
  get catalogDictionary(): string { return this.root === undefined ? "" : " /Outlines " + this.root + " 0 R"; }
  async finish(): Promise<void> {
    if (this.root === undefined) return;
    const descendants = new Map<number, number>();
    for (const item of [...this.items].reverse()) descendants.set(item.id, item.children.reduce((n, child) => n + 1 + descendants.get(child.id)!, 0));
    for (const item of this.items) {
      checkAbort(this.settings.options.signal);
      await this.objects.object(item.id, "<< /Title <feff" + unicodeHex(item.title) + "> /Parent " + item.parent + " 0 R" +
        (item.previous === undefined ? "" : " /Prev " + item.previous + " 0 R") + (item.next === undefined ? "" : " /Next " + item.next + " 0 R") +
        (item.children.length ? " /First " + item.children[0]!.id + " 0 R /Last " + item.children.at(-1)!.id + " 0 R /Count " + descendants.get(item.id) : "") +
        " /Dest [" + item.page + " 0 R /XYZ null " + pdfNumber(item.y) + " null] >>");
    }
    await this.objects.object(this.root, "<< /Type /Outlines /First " + this.top[0]!.id + " 0 R /Last " + this.top.at(-1)!.id + " 0 R /Count " + this.items.length + " >>");
  }
}
