import { OfficeIMOError } from "../core/errors.js";
import { checkAbort } from "../core/iteration.js";
import type { CellPresentation } from "../core/index.js";
import type { PdfObjects } from "./objects.js";
import { pdfNumber } from "./objects.js";
import { color } from "./settings.js";
import type { PdfSettings } from "./settings.js";
import type { PdfTableLayout, PdfLayoutRow } from "./table-layout.js";
import { wrapText, wrapTextCooperatively, synchronousText } from "./text.js";
import type { PdfFontResource, PdfFontResources } from "./font-resources.js";
import { linkAnnotation } from "./annotations.js";
import type { PdfStructure } from "./structure.js";
import { unicodeHex } from "./objects.js";

/** @internal One page of drawing commands and O(page count) references, never the complete report body. */
export class PdfPages {
  readonly references: number[] = [];
  private current = 0;
  private commands: string[] = [];
  private commandBytes = 0;
  private annotations: string[] = [];
  private hyperlinks = 0;
  private y = 0;
  private dataTop = 0;
  private continuationTop = 0;
  private readonly totalPages: number | undefined;
  private readonly pageNumberWidth: number;
  private readonly colors = new Map<string, string>();
  private readonly externalObjects = new Map<string, number>();
  private tableElement: number | undefined;
  private tableHeadingsDrawn = false;
  constructor(readonly objects: PdfObjects, readonly settings: PdfSettings, readonly fonts: PdfFontResources,
    private readonly parent: number, private readonly resources: number, private layout?: PdfTableLayout, private readonly structure?: PdfStructure) {
    this.totalPages = settings.options.pageNumbers === false ? undefined : objects.reserve();
    this.pageNumberWidth = 0;
    if (this.totalPages && settings.limits.maxPages > 0) {
      const font = fonts.select(), max = settings.limits.maxPages;
      const digitWidth = Math.max(...Array.from({ length: Math.min(10, max) }, (_, i) => font.width(max < 10 ? 49 + i : 48 + i)));
      const countWidth = digitWidth * String(max).length * settings.fontSize / 1000;
      const literals = wrapText("Page  of ", font, settings.fontSize, settings.page.width, 128, 1, false)[0]!.width;
      this.pageNumberWidth = literals + countWidth * 2 + 2;
      if (this.pageNumberWidth > settings.page.width - settings.margins.left - settings.margins.right)
        throw new RangeError("Page is too narrow for page numbers at the declared maxPages and font size.");
    }
  }
  private add(command: string): void {
    this.retainPageBytes(command.length);
    this.commands.push(command);
  }
  private retainPageBytes(length: number): void {
    this.commandBytes += length;
    if (this.commandBytes > this.settings.limits.maxPageBytes) throw new OfficeIMOError("RESOURCE_LIMIT", "maxPageBytes exceeded.");
  }
  private rgb(hex: string): string {
    let rgb = this.colors.get(hex);
    if (!rgb) { rgb = color(hex); if (this.colors.size < 1024) this.colors.set(hex, rgb); }
    return rgb;
  }
  private artifact(command: string, type = "Layout"): string {
    return this.structure ? "/Artifact << /Type /" + type + " >> BDC\n" + command + "EMC\n" : command;
  }
  private text(text: string, font: PdfFontResource, size: number, x: number, baseline: number, foreground = "111827", element?: number): void {
    const n = pdfNumber, shear = font.syntheticItalic ? .2126 : 0;
    const command = "q " + this.rgb(foreground) + " rg " + this.rgb(foreground) + " RG\nBT /" + font.name + " " + n(size) + " Tf " +
      (font.syntheticBold ? n(size * .025) + " w 2 Tr " : "0 Tr ") + "1 0 " + n(shear) + " 1 " + n(x) + " " + n(baseline) + " Tm " + font.encode(text) + " Tj ET Q\n";
    this.add(this.structure && element !== undefined ? this.structure.mark(element) + command + "EMC\n" : this.artifact(command, "Pagination"));
  }
  private decorationText(value: string | ((context: { pageNumber: number }) => string) | undefined): string {
    if (value === undefined) return "";
    checkAbort(this.settings.options.signal);
    return synchronousText(typeof value === "function" ? value({ pageNumber: this.references.length }) : value, "Page header/footer");
  }
  private decoration(text: string, x: number, y: number, width: number): void {
    if (!text) return;
    this.settings.budget.cell(text);
    const font = this.fonts.select(), lines = wrapText(text, font, this.settings.fontSize, width, this.settings.limits.maxCellCharacters, 1, false);
    this.text(lines[0]!.text, font, this.settings.fontSize, x, y);
  }
  async start(tableHeadings = true, preparedHeadings?: readonly PdfLayoutRow[]): Promise<void> {
    if (this.references.length >= this.settings.limits.maxPages) throw new OfficeIMOError("RESOURCE_LIMIT", "maxPages exceeded.");
    const { page, margins, options, fontSize } = this.settings;
    const lineHeight = this.fonts.select().lineHeight(fontSize);
    this.current = this.objects.reserve(); this.references.push(this.current); this.commands = []; this.commandBytes = 0;
    this.structure?.beginPage(this.current);
    this.y = page.height - margins.top;
    const header = this.decorationText(options.pageHeader), footer = this.decorationText(options.pageFooter);
    if (header && margins.top < lineHeight + 4) throw new RangeError("Top margin is too small for a page header.");
    if ((footer || this.totalPages) && margins.bottom < lineHeight + 4) throw new RangeError("Bottom margin is too small for page decorations.");
    const font = this.fonts.select(), centerOffset = ((font.program?.ascent ?? 800) + (font.program?.descent ?? -200)) * fontSize / 2000;
    this.decoration(header, margins.left, page.height - margins.top / 2 - centerOffset, page.width - margins.left - margins.right);
    const footerWidth = page.width - margins.left - margins.right - (this.totalPages ? this.pageNumberWidth + this.settings.padding * 2 : 0);
    if (footerWidth <= 0 && footer) throw new RangeError("Page is too narrow for both footer text and page numbers.");
    this.decoration(footer, margins.left, margins.bottom / 2 - centerOffset, footerWidth);
    if (this.totalPages) {
      const x = page.width - margins.right - this.pageNumberWidth + 1, y = margins.bottom / 2 - centerOffset, label = "Page " + this.references.length + " of ";
      const line = wrapText(label, font, fontSize, this.pageNumberWidth, 128, 1, false)[0]!;
      this.settings.budget.cell(label);
      this.text(label, font, fontSize, x, y);
      this.add(this.artifact("q 1 0 0 1 " + pdfNumber(x + line.width) + " " + pdfNumber(y) + " cm /TotalPages Do Q\n", "Pagination"));
    }
    if (this.references.length === 1) {
      if (options.title) await this.paragraph(options.title, { bold: true }, fontSize * 1.5, false, 1);
      if (options.messageTop) await this.paragraph(options.messageTop, {}, fontSize, false);
    }
    const headers = preparedHeadings ?? (tableHeadings && this.layout ? await this.layout.headers() : []), headerHeight = headers.reduce((n, h) => n + h.height, 0);
    this.continuationTop = page.height - margins.top - headerHeight;
    if (this.y - headerHeight - lineHeight - this.settings.padding * 2 < margins.bottom) throw new RangeError("PDF title and repeated headings leave no room for a data line.");
    if (tableHeadings && this.layout && this.structure && this.tableElement === undefined) this.tableElement = this.structure.element("Table");
    this.drawHeadings(headers);
    this.dataTop = this.y;
  }
  private drawHeadings(headers: readonly PdfLayoutRow[]): void {
    for (const header of headers) {
      const elements = this.structure && this.tableElement !== undefined && !this.tableHeadingsDrawn ?
        this.structure.tableRow(this.tableElement, header.cells.map(cell => cell.lines.some(line => line.text.trim().length > 0))) : undefined;
      this.draw(header, 0, header.lines, elements);
    }
    if (headers.length) this.tableHeadingsDrawn = true;
  }
  private rowElements(row: PdfLayoutRow): readonly number[] | undefined {
    return this.structure && this.tableElement !== undefined ? this.structure.tableRow(this.tableElement, row.cells.map(() => false)) : undefined;
  }
  private draw(row: PdfLayoutRow, firstLine: number, lineCount: number, elements?: readonly number[]): void {
    if (!this.layout) throw new Error("A table layout is required before drawing rows.");
    const { padding, fontSize } = this.settings, n = pdfNumber, height = lineCount * row.lineHeight + padding * 2;
    for (const [index, cell] of row.cells.entries()) {
      const x = this.layout.lefts[cell.first]!, cellHeight = cell.height ?? height;
      if (cell.style.background) this.add(this.artifact("q " + this.rgb(cell.style.background) + " rg " + n(x) + " " + n(this.y - cellHeight) + " " + n(cell.width) + " " + n(cellHeight) + " re f Q\n"));
      this.add(this.artifact("q 0.82 0.85 0.89 RG 0.4 w " + n(x) + " " + n(this.y - cellHeight) + " " + n(cell.width) + " " + n(cellHeight) + " re S Q\n"));
      if (cell.link) {
        if (++this.hyperlinks > this.settings.limits.maxHyperlinks) throw new OfficeIMOError("RESOURCE_LIMIT", "maxHyperlinks exceeded.");
        const annotation = linkAnnotation(cell.link, x, this.y - cellHeight, cell.width, cellHeight,
          this.settings.limits.maxPageBytes - this.commandBytes);
        this.retainPageBytes(annotation.length);
        this.annotations.push(annotation);
      }
      for (let line = firstLine; line < Math.min(cell.lines.length, cell.height ? cell.lines.length : firstLine + lineCount); line++) {
        const text = cell.lines[line]!, left = cell.style.alignment === "right" ? cell.width - padding - text.width : cell.style.alignment === "center" ? (cell.width - text.width) / 2 : padding;
        const ascent = (cell.font.program?.ascent ?? 800) * fontSize / 1000;
        this.text(text.text, cell.font, fontSize, x + left, this.y - padding - ascent - (line - firstLine) * (row.structured ? cell.lineHeight : row.lineHeight), cell.style.color, elements?.[index]);
      }
    }
    this.y -= row.structured ? row.height : height;
  }
  async block(rows: readonly PdfLayoutRow[]): Promise<void> {
    const height = rows.reduce((n, row) => n + row.height, 0), bottom = this.settings.margins.bottom;
    if (height > this.y - bottom) { await this.finishPage(); await this.start(); }
    if (height > this.dataTop - bottom) throw new RangeError("A structured footer must fit on one page below the repeated headings.");
    for (const row of rows) this.draw(row, 0, row.lines, this.rowElements(row));
  }
  async row(row: PdfLayoutRow): Promise<void> {
    const elements = this.rowElements(row);
    const { margins, padding } = this.settings, full = this.continuationTop - margins.bottom;
    if (row.height <= full + .001 && row.height > this.y - margins.bottom + .001) { await this.finishPage(); await this.start(); }
    let first = 0;
    while (first < row.lines) {
      let lines = Math.min(row.lines - first, Math.floor((this.y - margins.bottom - padding * 2 + .001) / row.lineHeight));
      if (lines <= 0) { await this.finishPage(); await this.start(); lines = Math.min(row.lines - first, Math.floor((this.y - margins.bottom - padding * 2 + .001) / row.lineHeight)); }
      if (lines <= 0) throw new RangeError("PDF page cannot fit a table line.");
      this.draw(row, first, lines, elements); first += lines;
      if (first < row.lines) { await this.finishPage(); await this.start(); }
    }
  }
  async paragraph(text: string, style: CellPresentation = {}, size = this.settings.fontSize, paginate = true, headingLevel?: number): Promise<void> {
    const { margins, page, limits } = this.settings, font = this.fonts.select(style.bold, style.italic), lineHeight = font.lineHeight(size);
    this.settings.budget.cell(text);
    const lines = await wrapTextCooperatively(text, font, size, page.width - margins.left - margins.right,
      limits.maxCellCharacters, limits.maxRowLines, true, this.settings.options.signal);
    if (lineHeight > page.height - margins.top - margins.bottom) throw new RangeError("PDF page cannot fit a paragraph line at the selected font metrics.");
    if (!paginate && this.y - lines.length * lineHeight - 8 < margins.bottom) throw new RangeError("PDF title/message exceeds the first page's content area.");
    const element = this.structure?.element(headingLevel ? "H" + headingLevel : "P");
    for (const line of lines) {
      if (this.y - lineHeight < margins.bottom) { await this.finishPage(); await this.start(false); }
      const x = style.alignment === "right" ? page.width - margins.right - line.width : style.alignment === "center" ? margins.left + (page.width - margins.left - margins.right - line.width) / 2 : margins.left;
      this.text(line.text, font, size, x, this.y - (font.program?.ascent ?? 800) * size / 1000, style.color, element);
      this.y -= lineHeight;
    }
    this.y -= 8;
  }
  /** Begin an independent table on the current page, retaining its headings on continuation pages. */
  async beginTable(layout: PdfTableLayout): Promise<void> {
    this.layout = layout;
    this.tableElement = this.structure?.element("Table"); this.tableHeadingsDrawn = false;
    const headings = await layout.headers(), height = headings.reduce((n, row) => n + row.height, 0);
    if (this.y - height - layout.lineHeight - this.settings.padding * 2 < this.settings.margins.bottom) {
      await this.finishPage(); await this.start(true, headings); return;
    }
    this.continuationTop = this.settings.page.height - this.settings.margins.top - height;
    this.drawHeadings(headings);
    this.dataTop = this.y;
  }
  endTable(): void { this.layout = undefined; this.tableElement = undefined; this.tableHeadingsDrawn = false; this.y -= 8; }
  async pageBreak(): Promise<void> { await this.finishPage(); await this.start(false); }
  registerXObject(name: string, id: number): void { this.externalObjects.set(name, id); }
  /** Place an indivisible graphic, preserving aspect ratio and scaling down to page bounds. */
  async graphic(width: number, height: number, alignment: "left" | "center" | "right",
    paint: (left: number, bottom: number, width: number, height: number) => string, alternativeText?: string): Promise<void> {
    const { page, margins } = this.settings, available = page.width - margins.left - margins.right;
    const scale = Math.min(1, available / width, (page.height - margins.top - margins.bottom) / height);
    width *= scale; height *= scale;
    if (this.y - height < margins.bottom) await this.pageBreak();
    const left = alignment === "right" ? page.width - margins.right - width : alignment === "center" ? margins.left + (available - width) / 2 : margins.left;
    const command = paint(left, this.y - height, width, height);
    const element = this.structure && alternativeText !== undefined ? this.structure.element("Figure", this.structure.document, " /Alt <feff" + unicodeHex(alternativeText) + ">") : undefined;
    this.add(this.structure && element !== undefined ? this.structure.mark(element) + command + "EMC\n" : this.artifact(command)); this.y -= height + 8;
  }
  private async finishPage(): Promise<void> {
    const contents = this.objects.reserve(), { page, options } = this.settings;
    await this.objects.stream(contents, this.commands.join(""), "", options.compression !== false);
    const annotations: number[] = [];
    for (const annotation of this.annotations) {
      const id = this.objects.reserve(); annotations.push(id);
      await this.objects.object(id, annotation);
    }
    await this.objects.object(this.current, "<< /Type /Page /Parent " + this.parent + " 0 R /MediaBox [0 0 " + pdfNumber(page.width) + " " + pdfNumber(page.height) +
      "] /Resources " + this.resources + " 0 R /Contents " + contents + " 0 R" +
      (this.structure?.pageDictionary ?? "") +
      (annotations.length ? " /Annots [" + annotations.map(id => id + " 0 R").join(" ") + "]" : "") + " >>");
    if (this.structure) await this.structure.finishPage();
    this.commands = []; this.annotations = []; this.commandBytes = 0;
    await this.objects.flush();
  }
  async finish(): Promise<void> {
    await this.finishPage();
    if (this.totalPages) {
      const font = this.fonts.select(), text = String(this.references.length), size = this.settings.fontSize;
      this.settings.budget.cell(text);
      const content = "BT /" + font.name + " " + pdfNumber(size) + " Tf 1 0 0 1 0 0 Tm " + font.encode(text) + " Tj ET\n";
      const width = wrapText(text, font, size, this.settings.page.width, 128, 1, false)[0]!.width;
      const bottom = (font.program?.bbox[1] ?? -250) * size / 1000 - 1, top = (font.program?.bbox[3] ?? 1000) * size / 1000 + 1;
      await this.objects.stream(this.totalPages, content, "/Type /XObject /Subtype /Form /BBox [-1 " + pdfNumber(bottom) + " " + pdfNumber(width + 1) + " " + pdfNumber(top) + "] /Resources << /Font " + this.fonts.dictionary() + " >>", this.settings.options.compression !== false);
    }
  }
  xobjects(): string {
    const entries = [...this.externalObjects].map(([name, id]) => "/" + name + " " + id + " 0 R");
    if (this.totalPages) entries.push("/TotalPages " + this.totalPages + " 0 R");
    return entries.length ? " /XObject << " + entries.join(" ") + " >>" : "";
  }
}
