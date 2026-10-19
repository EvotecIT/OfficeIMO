import type { ByteSink } from "../core/index.js";
import { PdfObjects, unicodeHex } from "./objects.js";
import { PdfFontResources } from "./font-resources.js";
import { PdfTableLayout } from "./table-layout.js";
import { PdfPages } from "./pages.js";
import type { PdfSettings } from "./settings.js";

/** One resource/catalog owner shared by table convenience exports and composed documents. */
export class PdfWriteSession {
  readonly objects: PdfObjects;
  readonly fonts: PdfFontResources;
  readonly pages: PdfPages;
  readonly layout: PdfTableLayout | undefined;
  private readonly root: number;
  private readonly parent: number;
  private readonly resources: number;
  private readonly info: number;
  constructor(sink: ByteSink, private readonly settings: PdfSettings, table = false) {
    this.objects = new PdfObjects(sink, settings.options.signal, settings.options.limits?.maxOutputBytes);
    this.root = this.objects.reserve(); this.parent = this.objects.reserve(); this.resources = this.objects.reserve(); this.info = this.objects.reserve();
    this.fonts = new PdfFontResources(this.objects, settings.options.fonts, settings.limits.maxFontBytes);
    this.layout = table ? new PdfTableLayout(settings, this.fonts) : undefined;
    this.pages = new PdfPages(this.objects, settings, this.fonts, this.parent, this.resources, this.layout);
  }
  async start(): Promise<void> {
    await this.objects.text("%PDF-1.7\n"); await this.objects.raw(Uint8Array.of(0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a));
    await this.pages.start(this.layout !== undefined);
  }
  async finish(): Promise<void> {
    const { options } = this.settings, { objects, fonts, pages } = this;
    await pages.finish();
    for (const font of fonts.resources) await font.write(objects, options.compression !== false);
    await objects.object(this.resources, "<< /Font " + fonts.dictionary() + pages.xobjects() + " >>");
    await objects.object(this.parent, "<< /Type /Pages /Count " + pages.references.length + " /Kids [" + pages.references.map(id => id + " 0 R").join(" ") + "] >>");
    await objects.object(this.root, "<< /Type /Catalog /Pages " + this.parent + " 0 R >>");
    await objects.object(this.info, "<< /Producer (OfficeIMO.js)" + (options.title === undefined ? "" : " /Title <feff" + unicodeHex(options.title) + ">") + " >>");
    await objects.finish(this.root, this.info);
  }
}
