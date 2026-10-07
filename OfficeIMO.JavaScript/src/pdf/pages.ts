import { OfficeIMOError } from "../core/errors.js";
import { checkAbort } from "../core/iteration.js";
import type { CellPresentation } from "../core/index.js";
import type { PdfObjects } from "./objects.js";
import { pdfNumber } from "./objects.js";
import { color } from "./settings.js";
import type { PdfSettings } from "./settings.js";
import type { PdfTableLayout, PdfLayoutRow } from "./table-layout.js";
import { wrapText, synchronousText } from "./text.js";
import type { PdfFontResource } from "./font-resources.js";

/** @internal One page of drawing commands and O(page count) references, never the complete report body. */
export class PdfPages {
  readonly references: number[] = [];
  private current = 0;
  private commands: string[] = [];
  private commandBytes = 0;
  private y = 0;
  private dataTop = 0;
  private readonly totalPages: number | undefined;
  private readonly colors = new Map<string, string>();
  constructor(readonly objects: PdfObjects, readonly settings: PdfSettings, readonly layout: PdfTableLayout,
    private readonly parent: number, private readonly resources: number) {
    this.totalPages = settings.options.pageNumbers === false ? undefined : objects.reserve();
  }
  private add(command: string): void {
    this.commandBytes += command.length;
    if (this.commandBytes > this.settings.limits.maxPageBytes) throw new OfficeIMOError("RESOURCE_LIMIT", "maxPageBytes exceeded.");
    this.commands.push(command);
  }
  private rgb(hex: string): string {
    let rgb = this.colors.get(hex);
    if (!rgb) { rgb = color(hex); if (this.colors.size < 1024) this.colors.set(hex, rgb); }
    return rgb;
  }
  private text(text: string, font: PdfFontResource, size: number, x: number, baseline: number, foreground = "111827"): void {
    const n = pdfNumber, shear = font.syntheticItalic ? .2126 : 0;
    this.add("q " + this.rgb(foreground) + " rg " + this.rgb(foreground) + " RG\nBT /" + font.name + " " + n(size) + " Tf " +
      (font.syntheticBold ? n(size * .025) + " w 2 Tr " : "0 Tr ") + "1 0 " + n(shear) + " 1 " + n(x) + " " + n(baseline) + " Tm " + font.encode(text) + " Tj ET Q\n");
  }
  private decoration(value: string | ((context: { pageNumber: number }) => string) | undefined, x: number, y: number, width: number): void {
    if (value === undefined) return;
    checkAbort(this.settings.options.signal);
    const text = synchronousText(typeof value === "function" ? value({ pageNumber: this.references.length }) : value, "Page header/footer");
    this.settings.budget.cell(text);
    const font = this.layout.fonts.select(), lines = wrapText(text, font, this.settings.fontSize, width, this.settings.limits.maxCellCharacters, 1, false);
    this.text(lines[0]!.text, font, this.settings.fontSize, x, y);
  }
  async start(): Promise<void> {
    if (this.references.length >= this.settings.limits.maxPages) throw new OfficeIMOError("RESOURCE_LIMIT", "maxPages exceeded.");
    const { page, margins, options, fontSize } = this.settings;
    this.current = this.objects.reserve(); this.references.push(this.current); this.commands = []; this.commandBytes = 0;
    this.y = page.height - margins.top;
    if (options.pageHeader !== undefined && margins.top < this.layout.lineHeight + 4) throw new RangeError("Top margin is too small for a page header.");
    if ((options.pageFooter !== undefined || this.totalPages) && margins.bottom < this.layout.lineHeight + 4) throw new RangeError("Bottom margin is too small for page decorations.");
    this.decoration(options.pageHeader, margins.left, page.height - margins.top / 2 - fontSize / 3, page.width - margins.left - margins.right);
    const footerWidth = page.width - margins.left - margins.right - (this.totalPages ? 112 : 0);
    if (footerWidth <= 0 && options.pageFooter !== undefined) throw new RangeError("Page is too narrow for both footer text and page numbers.");
    this.decoration(options.pageFooter, margins.left, margins.bottom / 2 - fontSize / 3, footerWidth);
    if (this.totalPages) {
      const x = page.width - margins.right - 110, y = margins.bottom / 2 - fontSize / 3, label = "Page " + this.references.length + " of ", font = this.layout.fonts.select();
      const line = wrapText(label, font, fontSize, 95, 128, 1, false)[0]!;
      this.settings.budget.cell(label);
      this.text(label, font, fontSize, x, y);
      this.add("q 1 0 0 1 " + pdfNumber(x + line.width) + " " + pdfNumber(y) + " cm /TotalPages Do Q\n");
    }
    if (this.references.length === 1) {
      if (options.title !== undefined) await this.paragraph(options.title, { bold: true }, fontSize * 1.5, false);
      if (options.messageTop !== undefined) await this.paragraph(options.messageTop, {}, fontSize, false);
    }
    const headers = this.layout.headers(), headerHeight = headers.reduce((n, h) => n + h.height, 0);
    if (this.y - headerHeight - this.layout.lineHeight - this.settings.padding * 2 < margins.bottom) throw new RangeError("PDF title and repeated headings leave no room for a data line.");
    for (const header of headers) this.draw(header, 0, header.lines);
    this.dataTop = this.y;
  }
  private draw(row: PdfLayoutRow, firstLine: number, lineCount: number): void {
    const { padding, fontSize } = this.settings, n = pdfNumber, height = lineCount * this.layout.lineHeight + padding * 2;
    for (const cell of row.cells) {
      const x = this.layout.lefts[cell.first]!, cellHeight = cell.height ?? height;
      if (cell.style.background) this.add("q " + this.rgb(cell.style.background) + " rg " + n(x) + " " + n(this.y - cellHeight) + " " + n(cell.width) + " " + n(cellHeight) + " re f Q\n");
      this.add("q 0.82 0.85 0.89 RG 0.4 w " + n(x) + " " + n(this.y - cellHeight) + " " + n(cell.width) + " " + n(cellHeight) + " re S Q\n");
      for (let line = firstLine; line < Math.min(cell.lines.length, cell.height ? cell.lines.length : firstLine + lineCount); line++) {
        const text = cell.lines[line]!, left = cell.style.alignment === "right" ? cell.width - padding - text.width : cell.style.alignment === "center" ? (cell.width - text.width) / 2 : padding;
        const ascent = (cell.font.program?.ascent ?? 800) * fontSize / 1000;
        this.text(text.text, cell.font, fontSize, x + left, this.y - padding - ascent - (line - firstLine) * this.layout.lineHeight, cell.style.color);
      }
    }
    this.y -= row.structured ? row.height : height;
  }
  async block(rows: readonly PdfLayoutRow[]): Promise<void> {
    const height = rows.reduce((n, row) => n + row.height, 0), bottom = this.settings.margins.bottom;
    if (height > this.y - bottom) { await this.finishPage(); await this.start(); }
    if (height > this.dataTop - bottom) throw new RangeError("A structured footer must fit on one page below the repeated headings.");
    for (const row of rows) this.draw(row, 0, row.lines);
  }
  async row(row: PdfLayoutRow): Promise<void> {
    const { margins, padding } = this.settings, full = this.dataTop - margins.bottom;
    if (row.height <= full + .001 && row.height > this.y - margins.bottom + .001 && this.y < this.dataTop - .001) { await this.finishPage(); await this.start(); }
    let first = 0;
    while (first < row.lines) {
      let lines = Math.min(row.lines - first, Math.floor((this.y - margins.bottom - padding * 2 + .001) / this.layout.lineHeight));
      if (lines <= 0) { await this.finishPage(); await this.start(); lines = Math.min(row.lines - first, Math.floor((this.y - margins.bottom - padding * 2 + .001) / this.layout.lineHeight)); }
      if (lines <= 0) throw new RangeError("PDF page cannot fit a table line.");
      this.draw(row, first, lines); first += lines;
      if (first < row.lines) { await this.finishPage(); await this.start(); }
    }
  }
  async paragraph(text: string, style: CellPresentation = {}, size = this.settings.fontSize, paginate = true): Promise<void> {
    const { margins, page, limits } = this.settings, font = this.layout.fonts.select(style.bold, style.italic), lineHeight = this.layout.lineHeight * size / this.settings.fontSize;
    this.settings.budget.cell(text);
    const lines = wrapText(text, font, size, page.width - margins.left - margins.right, limits.maxCellCharacters, limits.maxRowLines);
    if (!paginate && this.y - lines.length * lineHeight - 8 < margins.bottom) throw new RangeError("PDF title/message exceeds the first page's content area.");
    for (const line of lines) {
      if (this.y - lineHeight < margins.bottom) { await this.finishPage(); await this.start(); }
      this.text(line.text, font, size, margins.left, this.y - (font.program?.ascent ?? 800) * size / 1000, style.color);
      this.y -= lineHeight;
    }
    this.y -= 8;
  }
  private async finishPage(): Promise<void> {
    const contents = this.objects.reserve(), { page, options } = this.settings;
    await this.objects.stream(contents, this.commands.join(""), "", options.compression !== false);
    await this.objects.object(this.current, "<< /Type /Page /Parent " + this.parent + " 0 R /MediaBox [0 0 " + pdfNumber(page.width) + " " + pdfNumber(page.height) +
      "] /Resources " + this.resources + " 0 R /Contents " + contents + " 0 R >>");
    this.commands = []; this.commandBytes = 0;
    await this.objects.flush();
  }
  async finish(): Promise<void> {
    await this.finishPage();
    if (this.totalPages) {
      const font = this.layout.fonts.select(), text = String(this.references.length), size = this.settings.fontSize;
      this.settings.budget.cell(text);
      const content = "BT /" + font.name + " " + pdfNumber(size) + " Tf 1 0 0 1 0 0 Tm " + font.encode(text) + " Tj ET\n";
      const width = wrapText(text, font, size, this.settings.page.width, 128, 1, false)[0]!.width;
      const bottom = (font.program?.bbox[1] ?? -250) * size / 1000 - 1, top = (font.program?.bbox[3] ?? 1000) * size / 1000 + 1;
      await this.objects.stream(this.totalPages, content, "/Type /XObject /Subtype /Form /BBox [-1 " + pdfNumber(bottom) + " " + pdfNumber(width + 1) + " " + pdfNumber(top) + "] /Resources << /Font " + this.layout.fonts.dictionary() + " >>", this.settings.options.compression !== false);
    }
  }
  xobjects(): string { return this.totalPages ? " /XObject << /TotalPages " + this.totalPages + " 0 R >>" : ""; }
}
