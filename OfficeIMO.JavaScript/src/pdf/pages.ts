import { OfficeIMOError } from "../core/errors.js";
import { checkAbort } from "../core/iteration.js";
import type { CellPresentation } from "../core/index.js";
import type { ExportLink } from "../core/index.js";
import type { PdfObjects } from "./objects.js";
import { pdfNumber } from "./objects.js";
import { color } from "./settings.js";
import type { PdfSettings } from "./settings.js";
import { PdfTableBandFlow } from "./table-band-flow.js";
import type { PdfTableLayout, PdfLayoutRow, PdfLayoutCell } from "./table-layout.js";
import { PdfTableStructure } from "./table-structure.js";
import { wrapText, wrapTextCooperatively, synchronousText } from "./text.js";
import type { PdfFontResource, PdfFontResources } from "./font-resources.js";
import { linkAnnotation } from "./annotations.js";
import type { PdfStructure } from "./structure.js";
import { unicodeHex } from "./objects.js";
import type { PdfInlineLine } from "./inline-text.js";
import type { PdfPreparedParagraph } from "./paragraph-layout.js";
import type { PdfPreparedListItem } from "./list-layout.js";
import type { PdfSectionBlock, PdfBookmark } from "./section-types.js";
import { captureSection, captureBookmark } from "./section-types.js";
import { sectionSettings } from "./page-context.js";
import type { PdfOutlines } from "./outlines.js";
import type { PdfPageContext } from "./types.js";

interface PdfPaintTableLayout { readonly lefts: readonly number[]; readonly lineHeight: number; headers(): Promise<readonly PdfLayoutRow[]>; }

interface ParagraphFrame {
  readonly parent?: number;
  readonly role?: string;
  readonly indent?: number;
  readonly marker?: { readonly line: PdfInlineLine; readonly owner?: number; readonly gap: number };
  readonly artifact?: boolean;
  readonly owner?: number;
  readonly annotationParent?: number;
  readonly keepOpen?: boolean;
}

/** @internal One page of drawing commands and O(page count) references, never the complete report body. */
export class PdfPages {
  readonly references: number[] = [];
  private current = 0;
  private commands: string[] = [];
  private commandBytes = 0;
  private annotations: { readonly id: number; readonly body: string }[] = [];
  private hyperlinks = 0;
  private y = 0;
  private dataTop = 0;
  private continuationTop = 0;
  private totalPages: number | undefined;
  private numberSettings: PdfSettings | undefined;
  private activeSettings: PdfSettings;
  private pendingPage: number | undefined;
  private bodyAdded = false;
  private sectionCount = 0;
  private sectionPages = 0;
  private sectionTitle: string | undefined;
  private readonly pendingBookmarks: Readonly<PdfBookmark>[] = [];
  private pageNumberWidth = 0;
  private readonly colors = new Map<string, string>();
  private readonly externalObjects = new Map<string, number>();
  private bandFlow: PdfTableBandFlow | undefined;
  private bandCellBudget = false;
  private tableElement: number | undefined;
  private tableStructure: PdfTableStructure | undefined;
  private activeHeadings: readonly PdfLayoutRow[] | undefined;
  private paintedHeaderCells = new WeakSet<PdfLayoutCell>();
  private readonly paintedHeaderBudgetCells = new WeakSet<PdfLayoutCell>();
  private readonly summaryRows = new WeakSet<PdfLayoutRow>();
  private readonly groups: { readonly level: number; readonly headingLevel: number; readonly element: number | undefined }[] = [];
  private continuationCaption: PdfPreparedParagraph | undefined;
  private continuationCaptionOwner: number | undefined;
  private continuationCaptionPaints = 0;
  constructor(readonly objects: PdfObjects, private readonly baseSettings: PdfSettings, readonly fonts: PdfFontResources,
    private readonly parent: number, private readonly resources: number, private layout?: PdfPaintTableLayout, private readonly structure?: PdfStructure, private readonly outlines?: PdfOutlines) {
    this.activeSettings = baseSettings;
  }
  private measurePageNumbers(): number {
    const settings = this.settings, font = this.fonts.select(), max = settings.limits.maxPages;
    const digitWidth = Math.max(...Array.from({ length: Math.min(10, max) }, (_, i) => font.width(max < 10 ? 49 + i : 48 + i)));
    const countWidth = digitWidth * String(max).length * settings.fontSize / 1000;
    const labels = settings.options.pageNumberLabels;
    const literals = wrapText((labels?.prefix ?? "Page ") + (labels?.separator ?? " of ") + (labels?.suffix ?? ""), font,
      settings.fontSize, Number.MAX_SAFE_INTEGER, settings.limits.maxCellCharacters, 1, false)[0]!.width;
    return literals + countWidth * 2 + 2;
  }
  get settings(): PdfSettings { return this.activeSettings; }
  async ensureStarted(): Promise<void> { if (!this.current) await this.start(false); }
  /** Section changes consume no rows and do not manufacture an empty transition page. */
  async beginSection(block: PdfSectionBlock): Promise<Readonly<PdfSectionBlock>> {
    const captured = captureSection(block, this.baseSettings), active = sectionSettings(this.baseSettings, captured);
    if (++this.sectionCount > this.baseSettings.limits.maxSections) throw new OfficeIMOError("RESOURCE_LIMIT", "maxSections exceeded.");
    if (this.current) {
      if (this.bodyAdded) await this.finishPage();
      else { this.pendingPage = this.current; this.current = 0; }
    }
    this.activeSettings = active; this.sectionPages = 0; this.sectionTitle = captured.title;
    await this.structure?.beginSection();
    return captured;
  }
  addBookmark(bookmark: PdfBookmark): void {
    const captured = captureBookmark(bookmark, this.settings)!;
    if ((this.outlines?.count ?? 0) + this.pendingBookmarks.length >= this.settings.limits.maxBookmarks)
      throw new OfficeIMOError("RESOURCE_LIMIT", "maxBookmarks exceeded.");
    this.settings.budget.cell(captured.title);
    if (this.current) this.outlines?.add(captured, this.current, this.y);
    else this.pendingBookmarks.push(captured);
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
    const n = pdfNumber, shear = font.syntheticItalic ? .213 : 0;
    const command = "q " + this.rgb(foreground) + " rg " + this.rgb(foreground) + " RG\nBT /" + font.name + " " + n(size) + " Tf " +
      (font.syntheticBold ? n(size * .025) + " w 2 Tr " : "0 Tr ") + "1 0 " + n(shear) + " 1 " + n(x) + " " + n(baseline) + " Tm " + font.encode(text) + " Tj ET Q\n";
    this.add(this.structure && element !== undefined ? this.structure.mark(element) + command + "EMC\n" : this.artifact(command, "Pagination"));
  }
  private link(link: ExportLink, x: number, bottom: number, width: number, height: number, element?: number): void {
    if (width <= 0) return;
    if (++this.hyperlinks > this.settings.limits.maxHyperlinks) throw new OfficeIMOError("RESOURCE_LIMIT", "maxHyperlinks exceeded.");
    const id = this.objects.reserve();
    const parentKey = this.structure && element !== undefined ? this.structure.annotation(element, id) : undefined;
    const annotation = linkAnnotation(link, x, bottom, width, height, this.settings.limits.maxPageBytes - this.commandBytes, parentKey, this.structure ? this.current : undefined);
    this.retainPageBytes(annotation.length); this.annotations.push({ id, body: annotation });
  }
  private inline(line: PdfInlineLine, size: number, x: number, baseline: number, element?: number, annotationParent = element): void {
    const lineActualText = !this.structure && line.actualText !== undefined && line.actualText.length > 0;
    if (lineActualText) this.add("/Span << /ActualText <FEFF" + unicodeHex(line.actualText!) + "> >> BDC\n");
    const marks = new Map<PdfInlineLine["segments"][number], { owner: number; mark: string }>();
    if (line.actualText !== undefined && this.structure && element !== undefined) {
      // Allocate MCIDs/Link children in logical order, then paint the independently positioned visual order.
      for (const segment of [...line.segments].sort((a, b) => (a.logicalStart ?? 0) - (b.logicalStart ?? 0))) {
        const owner = segment.link ? this.structure.link(element, segment.link) : element;
        marks.set(segment, { owner, mark: this.structure.mark(owner, segment.actualText) });
      }
    }
    for (const segment of line.segments) {
      const registered = marks.get(segment), owner = registered?.owner ?? (segment.link && this.structure && element !== undefined ? this.structure.link(element, segment.link) : element);
      if (segment.glyphs) {
        const font = segment.font, n = pdfNumber, shear = font.syntheticItalic ? .213 : 0;
        let cursor = x, command = "q " + this.rgb(segment.color ?? "111827") + " rg " + this.rgb(segment.color ?? "111827") + " RG\nBT /" + font.name + " " + n(size) + " Tf " +
          (font.syntheticBold ? n(size * .025) + " w 2 Tr " : "0 Tr ");
        for (const glyph of segment.glyphs) {
          command += "1 0 " + n(shear) + " 1 " + n(cursor + (glyph.xOffset + shear * glyph.yOffset) * size / 1000) + " " + n(baseline + glyph.yOffset * size / 1000) + " Tm " + font.encodeGlyph(glyph.id, glyph.logicalText) + " Tj\n";
          cursor += glyph.xAdvance * size / 1000;
        }
        command += "ET Q\n";
        this.add(this.structure && owner !== undefined ? (registered?.mark ?? this.structure.mark(owner, segment.actualText)) + command + "EMC\n" : this.artifact(command, "Pagination"));
      } else this.text(segment.text, segment.font, size, x, baseline, segment.color, owner);
      if (segment.link) {
        const ascent = (segment.font.program?.ascent ?? 800) * size / 1000;
        const descent = (segment.font.program?.descent ?? -200) * size / 1000;
        const annotationOwner = owner ?? (this.structure && annotationParent !== undefined ? this.structure.link(annotationParent, segment.link) : undefined);
        this.link(segment.link, x, baseline + descent, segment.width, ascent - descent, annotationOwner);
      }
      x += segment.width;
    }
    if (lineActualText) this.add("EMC\n");
  }
  private decorationText(value: string | ((context: PdfPageContext) => string) | undefined): string {
    if (value === undefined) return "";
    checkAbort(this.settings.options.signal);
    return synchronousText(typeof value === "function" ? value(Object.freeze({ pageNumber: this.references.length, sectionNumber: this.sectionCount || 1, sectionPageNumber: this.sectionPages, ...(this.sectionTitle === undefined ? {} : { sectionTitle: this.sectionTitle }) })) : value, "Page header/footer");
  }
  private decoration(text: string, x: number, y: number, width: number): void {
    if (!text) return;
    this.settings.budget.cell(text);
    const font = this.fonts.select(), lines = wrapText(text, font, this.settings.fontSize, width, this.settings.limits.maxCellCharacters, 1, false);
    this.text(lines[0]!.text, font, this.settings.fontSize, x, y);
  }
  async start(tableHeadings = true, preparedHeadings?: readonly PdfLayoutRow[]): Promise<void> {
    if (this.pendingPage === undefined && this.references.length >= this.settings.limits.maxPages) throw new OfficeIMOError("RESOURCE_LIMIT", "maxPages exceeded.");
    const { page, margins, options, fontSize } = this.settings;
    const lineHeight = this.fonts.select().lineHeight(fontSize);
    if (this.pendingPage !== undefined) { this.current = this.pendingPage; this.pendingPage = undefined; }
    else { this.current = this.objects.reserve(); this.references.push(this.current); }
    this.commands = []; this.annotations = []; this.commandBytes = 0; this.bodyAdded = false; this.sectionPages++;
    if (options.pageNumbers !== false) {
      this.totalPages ??= this.objects.reserve(); this.numberSettings ??= this.settings;
      if (!this.pageNumberWidth) this.pageNumberWidth = this.measurePageNumbers();
      if (this.pageNumberWidth > page.width - margins.left - margins.right)
        throw new RangeError("Page is too narrow for page numbers at the declared maxPages and font size.");
    }
    this.structure?.beginPage(this.current);
    this.y = page.height - margins.top;
    const header = this.decorationText(options.pageHeader), footer = this.decorationText(options.pageFooter);
    if (header && margins.top < lineHeight + 4) throw new RangeError("Top margin is too small for a page header.");
    if ((footer || options.pageNumbers !== false) && margins.bottom < lineHeight + 4) throw new RangeError("Bottom margin is too small for page decorations.");
    const font = this.fonts.select(), centerOffset = ((font.program?.ascent ?? 800) + (font.program?.descent ?? -200)) * fontSize / 2000;
    this.decoration(header, margins.left, page.height - margins.top / 2 - centerOffset, page.width - margins.left - margins.right);
    const footerWidth = page.width - margins.left - margins.right - (options.pageNumbers !== false ? this.pageNumberWidth + this.settings.padding * 2 : 0);
    if (footerWidth <= 0 && footer) throw new RangeError("Page is too narrow for both footer text and page numbers.");
    this.decoration(footer, margins.left, margins.bottom / 2 - centerOffset, footerWidth);
    if (options.pageNumbers !== false) {
      const labels = options.pageNumberLabels, suffix = labels?.suffix ?? "";
      const x = page.width - margins.right - this.pageNumberWidth + 1, y = margins.bottom / 2 - centerOffset,
        label = (labels?.prefix ?? "Page ") + this.references.length + (labels?.separator ?? " of ");
      const line = wrapText(label, font, fontSize, this.pageNumberWidth, this.settings.limits.maxCellCharacters, 1, false)[0]!;
      this.settings.budget.cell(label + suffix);
      this.text(label, font, fontSize, x, y);
      this.add(this.artifact("q 1 0 0 1 " + pdfNumber(x + line.width) + " " + pdfNumber(y) + " cm /TotalPages Do Q\n", "Pagination"));
    }
    for (const bookmark of this.pendingBookmarks.splice(0)) this.outlines?.add(bookmark, this.current, this.y);
    if (this.references.length === 1) {
      if (options.title) await this.paragraph(options.title, { bold: true }, fontSize * 1.5, false, 1);
      if (options.messageTop) await this.paragraph(options.messageTop, {}, fontSize, false);
    }
    const headers = preparedHeadings ?? (tableHeadings && this.layout ? this.activeHeadings ??= await this.layout.headers() : []), headerHeight = headers.reduce((n, h) => n + h.height, 0),
      continuation = tableHeadings && this.layout ? this.continuationCaption : undefined;
    if (continuation?.lines.length) {
      if (this.continuationCaptionPaints++ > 0) this.settings.budget.cell(continuation.text.text);
      this.continuationCaptionOwner ??= this.captionAnnotations(continuation);
      await this.paragraphs([continuation], headerHeight + lineHeight + this.settings.padding * 2, true,
        { artifact: true, ...(this.continuationCaptionOwner === undefined ? {} : { annotationParent: this.continuationCaptionOwner }) });
    }
    this.continuationTop = page.height - margins.top - headerHeight - (continuation?.height ?? 0);
    if (this.y - headerHeight - lineHeight - this.settings.padding * 2 < margins.bottom) throw new RangeError("PDF title and repeated headings leave no room for a data line.");
    if (tableHeadings && this.layout && this.structure && this.tableElement === undefined) {
      this.tableElement = this.structure.element("Table", this.groupParent);
      this.tableStructure = new PdfTableStructure(this.structure, this.tableElement, headers);
      await this.tableStructure.prepareHeaders();
    }
    await this.drawHeadings(headers);
    this.dataTop = this.y;
  }
  private async drawHeadings(headers: readonly PdfLayoutRow[]): Promise<void> {
    for (const header of headers) {
      // Prepared headings can be reused by continuation pages and several logical summary tables.
      // Charge physical repeats without recapturing callbacks or making them semantic duplicates.
      for (const cell of header.cells) if (!this.bandCellBudget && this.paintedHeaderBudgetCells.has(cell.source ?? cell))
        this.settings.budget.cell(cell.lines.map(line => line.text).join(""));
      const elements = this.tableStructure ? header.cells.map(cell => this.tableStructure!.owner(cell.source ?? cell)) : undefined;
      const repeated = header.cells.map(cell => this.paintedHeaderCells.has(cell.source ?? cell));
      this.draw(header, 0, header.lines, elements, repeated);
      for (const cell of header.cells) {
        this.paintedHeaderCells.add(cell.source ?? cell); this.paintedHeaderBudgetCells.add(cell.source ?? cell);
      }
    }
  }
  private async rowElements(row: PdfLayoutRow): Promise<readonly number[] | undefined> {
    const source = row.source ?? row;
    return this.tableStructure?.prepareRow(source, "data", this.tableElement, this.summaryRows.has(source) ? 0 : undefined);
  }
  private draw(row: PdfLayoutRow, firstLine: number, lineCount: number, elements?: readonly (number | undefined)[], repeated: readonly boolean[] = []): void {
    if (!this.layout) throw new Error("A table layout is required before drawing rows.");
    this.bodyAdded = true;
    const { padding, fontSize } = this.settings, n = pdfNumber, height = lineCount * row.lineHeight + padding * 2;
    for (const [index, cell] of row.cells.entries()) {
      if (this.bandCellBudget) this.settings.budget.cell((cell.height === undefined ? cell.lines.slice(firstLine, firstLine + lineCount) : cell.lines).map(line => line.text).join(""));
      const x = this.layout.lefts[cell.first]!, cellHeight = cell.height ?? height;
      if (cell.style.background) this.add(this.artifact("q " + this.rgb(cell.style.background) + " rg " + n(x) + " " + n(this.y - cellHeight) + " " + n(cell.width) + " " + n(cellHeight) + " re f Q\n"));
      this.add(this.artifact("q 0.82 0.85 0.89 RG 0.4 w " + n(x) + " " + n(this.y - cellHeight) + " " + n(cell.width) + " " + n(cellHeight) + " re S Q\n"));
      const element = elements?.[index], artifact = cell.semanticArtifact || repeated[index],
        owner = cell.link && this.structure && element !== undefined ? this.structure.link(element, cell.link) : element;
      for (let line = firstLine; line < Math.min(cell.lines.length, cell.height ? cell.lines.length : firstLine + lineCount); line++) {
        const text = cell.lines[line]!, left = cell.style.alignment === "right" ? cell.width - padding - text.width : cell.style.alignment === "center" ? (cell.width - text.width) / 2 : padding;
        this.inline(text, fontSize, x + left, this.y - padding - text.ascent - (line - firstLine) * (row.structured ? cell.lineHeight : row.lineHeight), artifact ? undefined : owner, owner);
      }
      if (cell.link) this.link(cell.link, x, this.y - cellHeight, cell.width, cellHeight, owner);
    }
    this.y -= row.structured ? row.height : height;
  }
  async block(rows: readonly PdfLayoutRow[]): Promise<void> {
    if (this.bandFlow) return this.bandFlow.block(rows);
    const height = rows.reduce((n, row) => n + row.height, 0), bottom = this.settings.margins.bottom;
    if (height > this.y - bottom) { await this.finishPage(); await this.start(); }
    if (height > this.dataTop - bottom) throw new RangeError("A structured footer must fit on one page below the repeated headings.");
    for (const row of rows) {
      const elements = await this.rowElements(row); this.draw(row, 0, row.lines, elements);
      await this.tableStructure?.completeRow(row.source ?? row);
    }
  }
  async row(row: PdfLayoutRow, allowFirstFragment = false): Promise<void> {
    if (this.bandFlow) return this.bandFlow.row(row, allowFirstFragment);
    const elements = await this.rowElements(row);
    const { margins, padding } = this.settings, full = this.continuationTop - margins.bottom;
    if (!allowFirstFragment && row.height <= full + .001 && row.height > this.y - margins.bottom + .001) { await this.finishPage(); await this.start(); }
    let first = 0;
    while (first < row.lines) {
      let lines = Math.min(row.lines - first, Math.floor((this.y - margins.bottom - padding * 2 + .001) / row.lineHeight));
      if (lines <= 0) { await this.finishPage(); await this.start(); lines = Math.min(row.lines - first, Math.floor((this.y - margins.bottom - padding * 2 + .001) / row.lineHeight)); }
      if (lines <= 0) throw new RangeError("PDF page cannot fit a table line.");
      this.draw(row, first, lines, elements); first += lines;
      if (first < row.lines) { await this.finishPage(); await this.start(); }
    }
    await this.tableStructure?.completeRow(row.source ?? row);
  }
  /** Keep a last detail row and its closing subtotals together when the complete bundle fits below
   * fresh repeated context/headings. Oversized bundles deliberately use ordinary row splitting. */
  async keptRows(rows: readonly PdfLayoutRow[], allowFirstFragment = false): Promise<void> {
    if (this.bandFlow) return this.bandFlow.keptRows(rows, allowFirstFragment);
    const height = rows.reduce((total, row) => total + row.height, 0), bottom = this.settings.margins.bottom;
    const fits = height <= this.continuationTop - bottom + .001;
    if (fits && height > this.y - bottom + .001) { await this.finishPage(); await this.start(); }
    for (const [index, row] of rows.entries()) await this.row(row, !fits && index === 0 && allowFirstFragment);
  }
  async paragraph(text: string, style: CellPresentation = {}, size = this.settings.fontSize, paginate = true, headingLevel?: number): Promise<void> {
    await this.ensureStarted();
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
      this.bodyAdded = true;
      this.text(line.text, font, size, x, this.y - (font.program?.ascent ?? 800) * size / 1000, style.color, element);
      this.y -= lineHeight;
    }
    this.y -= 8;
    await this.structure?.complete(element);
  }
  /** Reserve a meaningful next fragment before painting a kept heading; never buffer a table. */
  async reserveSpace(height: number, label = "Kept PDF content"): Promise<void> {
    await this.ensureStarted();
    const { page, margins } = this.settings, full = page.height - margins.top - margins.bottom;
    if (height > full + .001) throw new RangeError(label + " exceeds a fresh page's content area.");
    if (height > this.y - margins.bottom + .001) await this.pageBreak();
  }
  /** Paint bounded paragraphs with their captured layout and a measured next-block fragment. */
  async paragraphs(paragraphs: readonly PdfPreparedParagraph[], nextHeight = 0, keepTogether = false, frame: ParagraphFrame = {}): Promise<void> {
    const visible = paragraphs.filter(paragraph => paragraph.lines.length > 0);
    if (!visible.length) return;
    await this.ensureStarted();
    const reservations = new Array<number>(visible.length).fill(0);
    for (let i = visible.length - 1; i >= 0; i--) {
      const paragraph = visible[i]!, next = visible[i + 1];
      reservations[i] = paragraph.keepWithNext ? next ? next.keepTogether || next.keepWithNext ? next.height + reservations[i + 1]! : next.firstHeight : nextHeight : 0;
    }
    if (keepTogether && visible.length) await this.reserveSpace(visible.reduce((height, paragraph) => height + paragraph.height, 0) + reservations.at(-1)!, "Kept PDF paragraph group");
    for (const [index, paragraph] of visible.entries()) {
      if (paragraph.keepTogether || paragraph.keepWithNext) await this.reserveSpace(paragraph.height + reservations[index]!, "Kept PDF paragraph chain");
      const { margins, page } = this.settings;
      if (paragraph.firstHeight > page.height - margins.top - margins.bottom + .001) throw new RangeError("PDF page cannot fit a paragraph line and its leading spacing.");
      if (this.y - paragraph.firstHeight < margins.bottom) await this.pageBreak();
      this.y -= paragraph.spaceBefore;
      if (paragraph.bookmark) this.addBookmark(paragraph.bookmark);
      const element = frame.artifact ? undefined : frame.owner ?? this.structure?.element(frame.role ?? (paragraph.headingLevel ? "H" + paragraph.headingLevel : "P"), frame.parent ?? this.groupParent);
      let firstLine = true;
      for (const line of paragraph.lines) {
        if (line.lineHeight > page.height - margins.top - margins.bottom + .001) throw new RangeError("PDF page cannot fit a paragraph line at the selected font metrics.");
        if (this.y - line.lineHeight < margins.bottom) await this.pageBreak();
        const left = margins.left + (frame.indent ?? 0), width = page.width - margins.right - left;
        const x = paragraph.style.alignment === "right" ? page.width - margins.right - line.width : paragraph.style.alignment === "center" ? left + (width - line.width) / 2 : left;
        if (firstLine && frame.marker) this.inline(frame.marker.line, paragraph.size, left - frame.marker.gap - frame.marker.line.width, this.y - line.ascent, frame.marker.owner);
        this.bodyAdded = true; this.inline(line, paragraph.size, x, this.y - line.ascent, element, frame.annotationParent ?? element); this.y -= line.lineHeight; firstLine = false;
      }
      this.y -= paragraph.spaceAfter;
      if (!frame.keepOpen) await this.structure?.complete(element);
    }
  }
  private get groupParent(): number | undefined { return this.groups.at(-1)?.element; }
  /** Nest a real report group around its visible heading, tables and closing summary. */
  async beginGroup(level: number, heading: PdfPreparedParagraph, minimumNextHeight: number): Promise<void> {
    if (level !== this.groups.length) throw new Error("PDF analytical groups must open in hierarchical order.");
    await this.ensureStarted();
    const element = this.structure?.element("Sect", this.groupParent);
    this.groups.push({ level, headingLevel: heading.headingLevel ?? level + 2, element });
    await this.paragraphs([heading], minimumNextHeight, true, { role: "H" + (heading.headingLevel ?? level + 2), ...(element === undefined ? {} : { parent: element }) });
  }
  async endGroup(level: number): Promise<void> {
    const current = this.groups.at(-1);
    if (!current || current.level !== level) throw new Error("PDF analytical groups must close in hierarchical order.");
    await this.structure?.complete(current.element); this.outlines?.closeGroup(current.headingLevel); this.groups.pop();
  }
  /** The generated summary label is a row heading; no localized text/style inference is used. */
  summary(row: PdfLayoutRow): void { this.summaryRows.add(row); }
  /** Begin an independent table on the current page, retaining its headings on continuation pages.
   * Group context uses captionArtifact: it stays visually complete without duplicating logical headings. */
  async beginTable(layout: PdfTableLayout, preparedHeadings?: readonly PdfLayoutRow[], minimumDataHeight = layout.lineHeight + this.settings.padding * 2, caption?: PdfPreparedParagraph,
    continuationCaption?: PdfPreparedParagraph, headingsOnNewPageOnly = false, captionArtifact = false): Promise<void> {
    if (layout.bands) return this.beginBands(layout, preparedHeadings, minimumDataHeight, caption, continuationCaption, captionArtifact);
    await this.ensureStarted();
    this.layout = layout;
    this.paintedHeaderCells = new WeakSet();
    this.continuationCaption = continuationCaption; this.continuationCaptionPaints = 0;
    this.continuationCaptionOwner = undefined;
    this.tableElement = this.structure?.element("Table", this.groupParent);
    const headings = preparedHeadings ?? await layout.headers();
    this.activeHeadings = headings;
    // A closing summary can share the preceding table's headings on the current page.
    // Its own headings remain available when it moves or splits onto a fresh page.
    const initialHeadings = headingsOnNewPageOnly && this.bodyAdded && !this.structure ? [] : headings;
    const height = initialHeadings.reduce((n, row) => n + row.height, 0);
    if (caption) {
      await this.reserveSpace(caption.height + height + minimumDataHeight, "Kept PDF table caption and first data fragment");
      const annotationParent = captionArtifact ? this.captionAnnotations(caption) : undefined;
      await this.paragraphs([caption], height + minimumDataHeight, true, captionArtifact
        ? { artifact: true, ...(annotationParent === undefined ? {} : { annotationParent }) }
        : { role: "Caption", ...(this.tableElement === undefined ? {} : { parent: this.tableElement }) });
    }
    if (this.structure) {
      this.tableStructure = new PdfTableStructure(this.structure, this.tableElement!, headings);
      await this.tableStructure.prepareHeaders();
    }
    if (this.y - height - minimumDataHeight < this.settings.margins.bottom) {
      await this.finishPage(); await this.start(true, headings); return;
    }
    this.continuationTop = this.settings.page.height - this.settings.margins.top - headings.reduce((n, row) => n + row.height, 0) - (continuationCaption?.height ?? 0);
    await this.drawHeadings(initialHeadings);
    this.dataTop = this.y;
  }
  /** Bind the bounded logical-window owner to the existing physical page painter. */
  private async beginBands(layout: PdfTableLayout, preparedHeadings: readonly PdfLayoutRow[] | undefined, minimumDataHeight: number,
    caption: PdfPreparedParagraph | undefined, continuation: PdfPreparedParagraph | undefined, captionArtifact: boolean): Promise<void> {
    await this.ensureStarted();
    const logicalHeadings = preparedHeadings ?? await layout.headers(), headings = await layout.bandRows(logicalHeadings, true);
    const headerHeight = headings[0]!.reduce((sum, row) => sum + row.height, 0), { page, margins, padding } = this.settings;
    const firstFullCapacity = page.height - margins.top - margins.bottom - headerHeight - (caption?.height ?? 0);
    const minimum = minimumDataHeight <= firstFullCapacity + .001 ? minimumDataHeight : layout.lineHeight + padding * 2;
    const continuationCapacity = page.height - margins.top - margins.bottom - headerHeight - (continuation?.height ?? 0);
    if (continuationCapacity + .001 < layout.lineHeight + padding * 2) throw new RangeError("PDF band headings and continuation context leave no room for a data line.");
    if (this.y - margins.bottom - headerHeight - (caption?.height ?? 0) + .001 < minimum) await this.pageBreak();
    let initialOffset = page.height - margins.top - this.y;
    const firstCapacity = this.y - margins.bottom - headerHeight - (caption?.height ?? 0);
    if (firstCapacity + .001 < minimum) throw new RangeError("PDF band caption and headings leave no room for a data line.");
    let captionPaints = 0, continuationPaints = 0, continuationOwner: number | undefined;
    this.tableElement = this.structure?.element("Table", this.groupParent);
    this.paintedHeaderCells = new WeakSet();
    const captionElement = this.structure && caption?.lines.length ? captionArtifact ? this.captionAnnotations(caption)
      : this.structure.element("Caption", this.tableElement) : undefined;
    if (this.structure) {
      this.tableStructure = new PdfTableStructure(this.structure, this.tableElement!, logicalHeadings);
    }
    this.layout = undefined; this.continuationCaption = undefined;
    this.activeHeadings = undefined;
    this.bandFlow = new PdfTableBandFlow(layout, headings, { firstCapacity, firstFullCapacity, continuationCapacity,
      prepareRow: async row => { await this.tableStructure?.prepareHeaders(); await this.rowElements(row); },
      completeRow: async row => { await this.tableStructure?.completeRow(row.source ?? row); },
      moveFirstToFreshPage: async () => { await this.pageBreak(); initialOffset = 0; },
      paint: async (band, headers, rows, currentCaption, firstWindow, firstBand) => {
        if (!(firstWindow && firstBand)) await this.pageBreak();
        if (firstWindow && !firstBand) this.y -= initialOffset;
        const frame = layout.bands![band]!;
        this.layout = { lefts: frame.lefts, lineHeight: layout.lineHeight, headers: async () => headers };
        this.bandCellBudget = true;
        try {
          if (currentCaption?.lines.length) {
            const paints = firstWindow ? captionPaints++ : continuationPaints++;
            if (paints > 0) this.settings.budget.cell(currentCaption.text.text);
            if (!firstWindow) continuationOwner ??= this.captionAnnotations(currentCaption);
            const annotationParent = firstWindow ? captionElement : continuationOwner;
            await this.paragraphs([currentCaption], headerHeight + minimum, true,
              firstWindow && firstBand && !captionArtifact ? { role: "Caption", keepOpen: true, ...(captionElement === undefined ? {} : { owner: captionElement }) }
                : { artifact: true, ...(annotationParent === undefined ? {} : { annotationParent }) });
          }
          await this.tableStructure?.prepareHeaders();
          await this.drawHeadings(headers); this.dataTop = this.y;
          for (const row of rows) {
            checkAbort(this.settings.options.signal);
            const elements = this.tableStructure ? row.cells.map(cell => this.tableStructure!.owner(cell.source ?? cell)) : undefined;
            this.draw(row, 0, row.lines, elements);
          }
        } finally { this.bandCellBudget = false; }
      }
    }, caption, continuation);
  }
  async endTable(): Promise<void> {
    if (this.bandFlow) { const flow = this.bandFlow; await flow.flush(true); this.bandFlow = undefined; }
    await this.tableStructure?.finish();
    this.layout = undefined; this.tableElement = undefined; this.tableStructure = undefined; this.activeHeadings = undefined; this.continuationCaption = undefined; this.continuationCaptionOwner = undefined; this.continuationCaptionPaints = 0; this.y -= 8;
  }
  /** Repeated context stays an artifact; its interactive links retain one authored label/owner. */
  private captionAnnotations(caption: PdfPreparedParagraph): number | undefined {
    if (!this.structure || !caption.text.runs.some(run => run.link && run.text.trim())) return undefined;
    const owner = this.structure.element("Caption", this.tableElement);
    for (const run of caption.text.runs) if (run.link && run.text.trim()) this.structure.link(owner, run.link, run.text);
    return owner;
  }
  async pageBreak(): Promise<void> { await this.ensureStarted(); await this.finishPage(); await this.start(false); }
  registerXObject(name: string, id: number): void { this.externalObjects.set(name, id); }
  /** Height after the existing fit-to-page graphic policy, used by heading keep constraints. */
  graphicHeight(width: number, height: number): number {
    const { page, margins } = this.settings;
    return height * Math.min(1, (page.width - margins.left - margins.right) / width, (page.height - margins.top - margins.bottom) / height);
  }
  /** Place an indivisible graphic, preserving aspect ratio and scaling down to page bounds. */
  async graphic(width: number, height: number, alignment: "left" | "center" | "right",
    paint: (left: number, bottom: number, width: number, height: number) => string, alternativeText?: string, caption?: PdfPreparedParagraph): Promise<void> {
    await this.ensureStarted();
    const { page, margins } = this.settings, available = page.width - margins.left - margins.right;
    const scale = Math.min(1, available / width, (page.height - margins.top - margins.bottom) / height);
    width *= scale; height *= scale;
    const element = this.structure && alternativeText !== undefined ? this.structure.element("Figure", undefined, " /Alt <feff" + unicodeHex(alternativeText) + ">") : undefined;
    if (caption) {
      if (alternativeText === undefined) throw new TypeError("A captioned PDF graphic requires alternativeText and cannot be decorative.");
      await this.reserveSpace(caption.height + height, "Kept PDF graphic and caption");
      await this.paragraphs([caption], height, true, { role: "Caption", ...(element === undefined ? {} : { parent: element }) });
    }
    if (this.y - height < margins.bottom) await this.pageBreak();
    const left = alignment === "right" ? page.width - margins.right - width : alignment === "center" ? margins.left + (available - width) / 2 : margins.left;
    const command = paint(left, this.y - height, width, height);
    this.bodyAdded = true;
    this.add(this.structure && element !== undefined ? this.structure.mark(element) + command + "EMC\n" : this.artifact(command)); this.y -= height + 8;
    await this.structure?.complete(element);
  }
  async listItem(item: PdfPreparedListItem, listOwner?: number): Promise<void> {
    await this.ensureStarted();
    if (this.structure && listOwner === undefined) throw new Error("A tagged list item requires a logical list owner.");
    const li = this.structure?.element("LI", listOwner), label = li === undefined ? undefined : this.structure!.element("Lbl", li),
      body = li === undefined ? undefined : this.structure!.element("LBody", li);
    await this.paragraphs([item.paragraph], 0, false, { ...(body === undefined ? {} : { parent: body }), role: "P", indent: item.indent,
      marker: { line: item.marker, gap: item.markerGap, ...(label === undefined ? {} : { owner: label }) } });
    await this.structure?.complete(li);
  }

  private async finishPage(): Promise<void> {
    const contents = this.objects.reserve(), { page, options } = this.settings;
    await this.objects.stream(contents, this.commands.join(""), "", options.compression !== false);
    const annotations: number[] = [];
    for (const annotation of this.annotations) {
      annotations.push(annotation.id);
      await this.objects.object(annotation.id, annotation.body);
    }
    await this.objects.object(this.current, "<< /Type /Page /Parent " + this.parent + " 0 R /MediaBox [0 0 " + pdfNumber(page.width) + " " + pdfNumber(page.height) +
      "] /Resources " + this.resources + " 0 R /Contents " + contents + " 0 R" +
      (this.structure?.pageDictionary ?? "") +
      (annotations.length ? " /Annots [" + annotations.map(id => id + " 0 R").join(" ") + "]" : "") + " >>");
    if (this.structure) await this.structure.finishPage();
    this.commands = []; this.annotations = []; this.commandBytes = 0; this.current = 0;
    await this.objects.flush();
  }
  async finish(): Promise<void> {
    if (this.bandFlow) { const flow = this.bandFlow; await flow.flush(true); this.bandFlow = undefined; }
    await this.ensureStarted();
    await this.tableStructure?.finish();
    await this.finishPage();
    if (this.totalPages) {
      const font = this.fonts.select(), count = String(this.references.length), text = count + (this.numberSettings!.options.pageNumberLabels?.suffix ?? ""), size = this.numberSettings!.fontSize;
      this.settings.budget.cell(count);
      const content = "BT /" + font.name + " " + pdfNumber(size) + " Tf 1 0 0 1 0 0 Tm " + font.encode(text) + " Tj ET\n";
      const width = wrapText(text, font, size, this.numberSettings!.page.width, this.settings.limits.maxCellCharacters, 1, false)[0]!.width;
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
