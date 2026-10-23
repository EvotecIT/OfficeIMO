import type { ByteSink } from "../core/index.js";
import { PdfObjects, unicodeHex } from "./objects.js";
import { PdfFontResources } from "./font-resources.js";
import { PdfTableLayout } from "./table-layout.js";
import { PdfPages } from "./pages.js";
import type { PdfSettings } from "./settings.js";
import { PdfStructure } from "./structure.js";
import { PdfOutlines } from "./outlines.js";

/** One resource/catalog owner shared by table convenience exports and composed documents. */
export class PdfWriteSession {
  readonly objects: PdfObjects;
  readonly fonts: PdfFontResources;
  readonly pages: PdfPages;
  readonly layout: PdfTableLayout | undefined;
  readonly structure: PdfStructure | undefined;
  readonly outlines: PdfOutlines;
  private readonly root: number;
  private readonly parent: number;
  private readonly resources: number;
  private readonly info: number;
  constructor(sink: ByteSink, private readonly settings: PdfSettings, table = false) {
    this.objects = new PdfObjects(sink, settings.options.signal, settings.budget.limits.maxOutputBytes, settings.options.compression !== false);
    this.root = this.objects.reserve(); this.parent = this.objects.reserve(); this.resources = this.objects.reserve(); this.info = this.objects.reserve();
    this.structure = settings.options.tagged ? new PdfStructure(this.objects, settings) : undefined;
    this.outlines = new PdfOutlines(this.objects, settings);
    this.fonts = new PdfFontResources(this.objects, settings.options.fonts, settings.limits.maxFontBytes, settings.options.signal);
    this.layout = table ? new PdfTableLayout(settings, this.fonts) : undefined;
    this.pages = new PdfPages(this.objects, settings, this.fonts, this.parent, this.resources, this.layout, this.structure, this.outlines);
  }
  async start(): Promise<void> {
    await this.objects.text("%PDF-1.7\n"); await this.objects.raw(Uint8Array.of(0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a));
    if (this.layout?.bands) await this.pages.beginTable(this.layout);
    else if (this.layout) await this.pages.start(true);
  }
  async finish(): Promise<void> {
    const { options } = this.settings, { objects, fonts, pages } = this;
    await pages.finish();
    for (const font of fonts.resources) await font.write(objects, options.compression !== false);
    await objects.object(this.resources, "<< /Font " + fonts.dictionary() + pages.xobjects() + " >>");
    await objects.object(this.parent, "<< /Type /Pages /Count " + pages.references.length + " /Kids [" + pages.references.map(id => id + " 0 R").join(" ") + "] >>");
    if (this.structure) await this.structure.finish();
    await this.outlines.finish();
    await objects.object(this.root, "<< /Type /Catalog /Pages " + this.parent + " 0 R" + (this.structure?.catalogDictionary ?? "") + this.outlines.catalogDictionary + " >>");
    await objects.object(this.info, "<< /Producer (OfficeIMO.js)" + (options.title === undefined ? "" : " /Title <feff" + unicodeHex(options.title) + ">") + " >>");
    await objects.finish(this.root, this.info);
  }
}
