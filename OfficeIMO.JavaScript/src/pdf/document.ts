import { BlobByteSink, withDestination } from "../core/sinks.js";
import { OfficeIMOError } from "../core/errors.js";
import { beginTask, checkAbort, consumeRows, pause, taskYieldDue } from "../core/iteration.js";
import { notifyProgress } from "../core/progress.js";
import type { ExportResult, ExportValue, OutputDestination } from "../core/index.js";
import { copyColumns, createRowProjector } from "../internal/rows.js";
import { settings } from "./settings.js";
import { pdfNumber as n } from "./objects.js";
import { PdfTableLayout } from "./table-layout.js";
import { PdfWriteSession } from "./session.js";
import { PdfDocumentImages } from "./document-images.js";
import { vectorCommands } from "./document-vectors.js";
import { graphicAlternative } from "./structure.js";
import { checkDrawingBlock, prepareDrawing } from "./document-drawings.js";
import type { PdfDocumentBlock, PdfDocumentOptions, PdfTableBlock, PdfTableOptions, PdfCaption } from "./document-types.js";
import { captureParagraph, captureCaption, prepareParagraph, checkParagraphBuffer, paragraphStartHeight } from "./paragraph-layout.js";
import type { PdfCapturedParagraph } from "./paragraph-layout.js";
import { captureListOptions, prepareListItem } from "./list-layout.js";
import { captureSection } from "./section-types.js";

/** Capture typed columns for a single-use table source in a composed document. */
export function pdfTable<T extends object>(rows: Iterable<T> | AsyncIterable<T>, options: PdfTableOptions<NoInfer<T>>): PdfTableBlock {
  checkTableOptions(options);
  return Object.freeze({ kind: "table", rows, options: Object.freeze({ ...options, columns: Object.freeze(copyColumns(options?.columns)) }) }) as PdfTableBlock;
}
/** Write sequential report blocks to a PDF Blob, retaining finished output proportional to file size. */
export async function writePdfDocument(blocks: Iterable<PdfDocumentBlock> | AsyncIterable<PdfDocumentBlock>, options: PdfDocumentOptions = {}): Promise<Blob> {
  const sink = new BlobByteSink();
  try { await writePdfDocumentTo(blocks, sink, options); return sink.toBlob("application/pdf"); }
  catch (error) { sink.discard(); throw error; }
}
/** Stream composed pages with backpressure. Rows sum all tables; columns reports the widest table. */
export async function writePdfDocumentTo(blocks: Iterable<PdfDocumentBlock> | AsyncIterable<PdfDocumentBlock>, destination: OutputDestination,
  configuration: PdfDocumentOptions = {}): Promise<ExportResult> {
  const prepared = settings({ ...configuration, columns: [], includeHeader: false }, true), { options, budget } = prepared;
  beginTask(); checkAbort(options.signal);
  return withDestination(destination, async sink => {
    const session = new PdfWriteSession(sink, prepared), { objects, pages, fonts } = session, images = new PdfDocumentImages(objects, prepared);
    let rows = 0, columns = 0, commands = 0, drawingElements = 0, listItems = 0;
    let pending: PdfCapturedParagraph[] = [];
    const flushParagraphs = async (nextHeight = 0): Promise<void> => {
      const paragraphs = [];
      for (const paragraph of pending) paragraphs.push(await prepareParagraph(paragraph, pages.settings, fonts));
      pending = [];
      await pages.paragraphs(paragraphs, nextHeight);
    };
    const enqueueParagraph = async (block: import("./document-types.js").PdfParagraphBlock): Promise<void> => {
      const paragraph = captureParagraph(block, pages.settings, fonts);
      if (!paragraph.text.text) return;
      if (pending.length || paragraph.keepWithNext) {
        pending.push(paragraph); checkParagraphBuffer(pending, pages.settings);
        if (!paragraph.keepWithNext) await flushParagraphs();
      } else await pages.paragraphs([await prepareParagraph(paragraph, pages.settings, fonts)]);
    };
    const prepareCaption = async (captured: PdfCapturedParagraph | undefined) =>
      captured === undefined ? undefined : prepareParagraph(captured, pages.settings, fonts);
    const snapshotCaption = (caption: PdfCaption | undefined) =>
      caption === undefined ? undefined : captureCaption(caption, pages.settings, fonts);
    await session.start();
    await consumeRows(blocks, options.signal, async block => {
      if (!block || typeof block !== "object") throw new TypeError("PDF document blocks must be objects.");
      switch (block.kind) {
        case "section": {
          const capturedSection = captureSection(block, prepared);
          await flushParagraphs();
          const section = await pages.beginSection(capturedSection);
          if (section.includeHeading) await enqueueParagraph({ kind: "paragraph", text: section.title,
            headingLevel: section.level!, ...(section.headingStyle === undefined ?
              { fontSize: pages.settings.fontSize * 1.5, presentation: { bold: true } } : { style: section.headingStyle }),
            keepTogether: true, keepWithNext: true,
            ...(section.bookmark ? { bookmark: { title: section.title, level: section.level! } } : {}) });
          else if (section.bookmark) pages.addBookmark({ title: section.title, level: section.level! });
          break;
        }
        case "paragraph": await enqueueParagraph(block); break;
        case "list": {
          const captured = captureListOptions(block, pages.settings);
          let count = 0, owner: number | undefined;
          await consumeRows(captured.items, options.signal, async item => {
            if (++listItems > prepared.limits.maxListItems) throw new OfficeIMOError("RESOURCE_LIMIT", "maxListItems exceeded.");
            const laidOut = await prepareListItem(item, count++, captured, pages.settings, fonts);
            if (count === 1) {
              await flushParagraphs(laidOut.paragraph.keepTogether ? laidOut.paragraph.height : laidOut.paragraph.firstHeight);
              owner = session.structure?.element("L");
            }
            await pages.listItem(laidOut, owner);
            if (taskYieldDue()) await pause(); checkAbort(options.signal);
          });
          await session.structure?.complete(owner); break;
        }
        case "group": {
          const blocks = block.blocks, keepTogether = block.keepTogether;
          if (!Array.isArray(blocks)) throw new TypeError("PDF paragraph group blocks must be an array.");
          if (keepTogether !== undefined && typeof keepTogether !== "boolean") throw new TypeError("PDF paragraph group keepTogether must be boolean.");
          if (blocks.length > prepared.limits.maxGroupBlocks) throw new OfficeIMOError("RESOURCE_LIMIT", "maxGroupBlocks exceeded.");
          const captured: PdfCapturedParagraph[] = [];
          for (const paragraph of blocks) { captured.push(captureParagraph(paragraph, pages.settings, fonts)); checkParagraphBuffer(captured, prepared); }
          if (captured.filter(paragraph => paragraph.text.text).at(-1)?.keepWithNext)
            throw new TypeError("The last paragraph in a group cannot keepWithNext beyond the group's boundary.");
          const paragraphs = [];
          for (const paragraph of captured) paragraphs.push(await prepareParagraph(paragraph, pages.settings, fonts));
          if (!paragraphs.some(paragraph => paragraph.lines.length > 0)) break;
          await flushParagraphs(keepTogether !== false ? paragraphs.reduce((height, paragraph) => height + paragraph.height, 0) : paragraphStartHeight(paragraphs));
          await pages.paragraphs(paragraphs, 0, keepTogether !== false); break;
        }
        case "pageBreak": await flushParagraphs(); await pages.pageBreak(); break;
        case "table": {
          const { options: tableOptions, rows: source, caption: inputCaption, continuationCaption: inputContinuationCaption } = block;
          checkTableOptions(tableOptions);
          const capturedCaption = snapshotCaption(inputCaption);
          if (inputContinuationCaption !== undefined && options.tagged)
            throw new TypeError("Tagged PDF table continuation captions are unsupported.");
          const capturedContinuationCaption = snapshotCaption(inputContinuationCaption);
          const table = settings({ ...pages.settings.options, ...tableOptions, columns: tableOptions?.columns, includeHeader: tableOptions?.includeHeader ?? true }), layout = new PdfTableLayout({ ...table, budget }, fonts);
          const caption = await prepareCaption(capturedCaption);
          const continuationCaption = await prepareCaption(capturedContinuationCaption);
          const project = createRowProjector(table.options.columns, undefined, options.signal);
          let count = 0, started = false; columns = Math.max(columns, table.options.columns.length);
          const headings = await layout.headers(), headingHeight = headings.reduce((height, row) => height + row.height, 0);
          await consumeRows(source, options.signal, async row => {
            budget.row(rows + 1);
            const laidOut = await layout.data(project(row, count) as readonly ExportValue[], count), keptHeading = !started && (pending.length > 0 || !!caption?.lines.length);
            if (!started) {
              const minimum = laidOut.lineHeight + table.padding * 2;
              await flushParagraphs((caption?.height ?? 0) + headingHeight + minimum);
              await pages.beginTable(layout, headings, minimum, caption, continuationCaption); started = true;
            }
            await pages.row(laidOut, keptHeading); count++; rows++;
            if (rows % 256 === 0) notifyProgress(options.onProgress, { phase: "rows", rows, bytes: objects.bytes });
            if (taskYieldDue()) await pause(); checkAbort(options.signal);
          });
          const footer = await layout.footer(count), keptFooterHeading = !started && (pending.length > 0 || !!caption?.lines.length);
          if (!started) {
            if (!headings.length && !footer?.length) break;
            const minimum = footer?.length ? table.options.footer?.rows ? footer.reduce((height, row) => height + row.height, 0) :
              footer[0]!.lineHeight + table.padding * 2 : 0;
            await flushParagraphs((caption?.height ?? 0) + headingHeight + minimum); await pages.beginTable(layout, headings, minimum, caption, continuationCaption);
          }
          if (footer) { if (table.options.footer?.rows) await pages.block(footer); else await pages.row(footer[0]!, keptFooterHeading); }
          await pages.endTable(); break;
        }
        case "image": {
          const { width: displayWidth, height: displayHeight, image: inputImage, caption: inputCaption } = block;
          const capturedCaption = snapshotCaption(inputCaption), alternative = graphicAlternative(block, pages.settings), alignment = graphicAlignment(block.alignment);
          if (capturedCaption && alternative === undefined) throw new TypeError("A captioned PDF graphic requires alternativeText and cannot be decorative.");
          const image = await images.prepare(inputImage), caption = await prepareCaption(capturedCaption);
          const width = displayWidth ?? (displayHeight === undefined ? image.width * .75 : displayHeight * image.width / image.height);
          const height = displayHeight ?? width * image.height / image.width;
          if (!Number.isFinite(width) || width <= 0 || !Number.isFinite(height) || height <= 0)
            throw new RangeError("Image display dimensions must be finite and positive.");
          await flushParagraphs((caption?.height ?? 0) + pages.graphicHeight(width, height));
          pages.registerXObject(image.name, image.id);
          await pages.graphic(width, height, alignment, (x, y, w, h) => "q " + [w, 0, 0, h, x, y].map(n).join(" ") + " cm /" + image.name + " Do Q\n", alternative, caption); break;
        }
        case "vector": {
          const capturedCaption = snapshotCaption(block.caption), alternative = graphicAlternative(block, pages.settings), alignment = graphicAlignment(block.alignment);
          if (capturedCaption && alternative === undefined) throw new TypeError("A captioned PDF graphic requires alternativeText and cannot be decorative.");
          const path = await vectorCommands(block, pages.settings, commands), caption = await prepareCaption(capturedCaption); commands += path.count;
          await flushParagraphs((caption?.height ?? 0) + pages.graphicHeight(path.width, path.height));
          await pages.graphic(path.width, path.height, alignment, (x, y, w, h) => "q " + [w / path.width, 0, 0, -h / path.height, x, y + h].map(n).join(" ") +
            " cm 0 0 " + n(path.width) + " " + n(path.height) + " re W n\n" + path.content + "Q\n", alternative, caption); break;
        }
        case "drawing": {
          checkDrawingBlock(block);
          const inputDrawing = block.drawing, capturedCaption = snapshotCaption(block.caption), alternative = graphicAlternative(block, pages.settings), alignment = graphicAlignment(block.alignment);
          if (capturedCaption && alternative === undefined) throw new TypeError("A captioned PDF graphic requires alternativeText and cannot be decorative.");
          const drawing = await prepareDrawing(inputDrawing, pages.settings, fonts, { commands, elements: drawingElements }), caption = await prepareCaption(capturedCaption);
          commands += drawing.commands; drawingElements += drawing.elements;
          await flushParagraphs((caption?.height ?? 0) + pages.graphicHeight(drawing.width, drawing.height));
          await pages.graphic(drawing.width, drawing.height, alignment, (x, y, w, h) => "q " + [w / drawing.width, 0, 0, -h / drawing.height, x, y + h].map(n).join(" ") +
            " cm 0 0 " + n(drawing.width) + " " + n(drawing.height) + " re W n\n" + drawing.content + "Q\n", alternative, caption); break;
        }
        default: throw new TypeError("Unsupported PDF document block.");
      }
      if (taskYieldDue()) await pause(); checkAbort(options.signal);
    });
    await flushParagraphs();
    if (options.messageBottom) await pages.paragraph(options.messageBottom);
    await session.finish();
    notifyProgress(options.onProgress, { phase: "complete", rows, bytes: objects.bytes }); checkAbort(options.signal);
    return { rows, columns, bytes: objects.bytes };
  });
}
function graphicAlignment(value: string | undefined): "left" | "center" | "right" {
  if (value !== undefined && !["left", "center", "right"].includes(value)) throw new TypeError("Invalid graphic alignment.");
  return value as "left" | "center" | "right" | undefined ?? "left";
}
/** Per-table overrides cannot change the document's shared fonts, geometry or resource policy. */
function checkTableOptions(options: unknown): void {
  const allowed = ["columns", "includeHeader", "headerRows", "footer", "headerPresentation", "footerPresentation", "alternateRowColor", "columnWidths", "wideTable", "formatValue", "formatRuns"];
  if (!options || typeof options !== "object" || Object.keys(options).some(key => !allowed.includes(key)))
    throw new TypeError("Table options support columns, headings, footers, widths and presentation; document settings belong on writePdfDocument.");
}
