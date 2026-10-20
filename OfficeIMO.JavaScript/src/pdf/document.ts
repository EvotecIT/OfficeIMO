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
import type { PdfDocumentBlock, PdfDocumentOptions, PdfTableBlock, PdfTableOptions } from "./document-types.js";
import { captureParagraph, prepareParagraph, checkParagraphBuffer, paragraphStartHeight } from "./paragraph-layout.js";
import type { PdfCapturedParagraph } from "./paragraph-layout.js";

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
    let rows = 0, columns = 0, commands = 0, drawingElements = 0;
    let pending: PdfCapturedParagraph[] = [];
    const flushParagraphs = async (nextHeight = 0): Promise<void> => {
      const paragraphs = [];
      for (const paragraph of pending) paragraphs.push(await prepareParagraph(paragraph, prepared, fonts));
      pending = [];
      await pages.paragraphs(paragraphs, nextHeight);
    };
    await session.start();
    await consumeRows(blocks, options.signal, async block => {
      if (!block || typeof block !== "object") throw new TypeError("PDF document blocks must be objects.");
      switch (block.kind) {
        case "paragraph": {
          const paragraph = captureParagraph(block, prepared, fonts);
          if (!paragraph.text.text) break;
          if (pending.length || paragraph.keepWithNext) {
            pending.push(paragraph); checkParagraphBuffer(pending, prepared);
            if (!paragraph.keepWithNext) await flushParagraphs();
          } else await pages.paragraphs([await prepareParagraph(paragraph, prepared, fonts)]);
          break;
        }
        case "group": {
          const blocks = block.blocks, keepTogether = block.keepTogether;
          if (!Array.isArray(blocks)) throw new TypeError("PDF paragraph group blocks must be an array.");
          if (keepTogether !== undefined && typeof keepTogether !== "boolean") throw new TypeError("PDF paragraph group keepTogether must be boolean.");
          if (blocks.length > prepared.limits.maxGroupBlocks) throw new OfficeIMOError("RESOURCE_LIMIT", "maxGroupBlocks exceeded.");
          const captured: PdfCapturedParagraph[] = [];
          for (const paragraph of blocks) { captured.push(captureParagraph(paragraph, prepared, fonts)); checkParagraphBuffer(captured, prepared); }
          if (captured.filter(paragraph => paragraph.text.text).at(-1)?.keepWithNext)
            throw new TypeError("The last paragraph in a group cannot keepWithNext beyond the group's boundary.");
          const paragraphs = [];
          for (const paragraph of captured) paragraphs.push(await prepareParagraph(paragraph, prepared, fonts));
          if (!paragraphs.some(paragraph => paragraph.lines.length > 0)) break;
          await flushParagraphs(keepTogether !== false ? paragraphs.reduce((height, paragraph) => height + paragraph.height, 0) : paragraphStartHeight(paragraphs));
          await pages.paragraphs(paragraphs, 0, keepTogether !== false); break;
        }
        case "pageBreak": await flushParagraphs(); await pages.pageBreak(); break;
        case "table": {
          checkTableOptions(block.options);
          const table = settings({ ...options, ...block.options, columns: block.options?.columns, includeHeader: block.options?.includeHeader ?? true }), layout = new PdfTableLayout({ ...table, budget }, fonts);
          const project = createRowProjector(table.options.columns, undefined, options.signal);
          let count = 0, started = false; columns = Math.max(columns, table.options.columns.length);
          const headings = await layout.headers(), headingHeight = headings.reduce((height, row) => height + row.height, 0);
          await consumeRows(block.rows, options.signal, async row => {
            budget.row(rows + 1);
            const laidOut = await layout.data(project(row, count) as readonly ExportValue[], count), keptHeading = !started && pending.length > 0;
            if (!started) {
              const minimum = laidOut.lineHeight + prepared.padding * 2;
              await flushParagraphs(headingHeight + minimum);
              await pages.beginTable(layout, headings, minimum); started = true;
            }
            await pages.row(laidOut, keptHeading); count++; rows++;
            if (rows % 256 === 0) notifyProgress(options.onProgress, { phase: "rows", rows, bytes: objects.bytes });
            if (taskYieldDue()) await pause(); checkAbort(options.signal);
          });
          const footer = await layout.footer(count), keptFooterHeading = !started && pending.length > 0;
          if (!started) {
            if (!headings.length && !footer?.length) break;
            const minimum = footer?.length ? table.options.footer?.rows ? footer.reduce((height, row) => height + row.height, 0) :
              footer[0]!.lineHeight + prepared.padding * 2 : 0;
            await flushParagraphs(headingHeight + minimum); await pages.beginTable(layout, headings, minimum);
          }
          if (footer) { if (table.options.footer?.rows) await pages.block(footer); else await pages.row(footer[0]!, keptFooterHeading); }
          pages.endTable(); break;
        }
        case "image": {
          const displayWidth = block.width, displayHeight = block.height;
          const alternative = graphicAlternative(block, prepared), alignment = graphicAlignment(block.alignment), image = await images.prepare(block.image);
          const width = displayWidth ?? (displayHeight === undefined ? image.width * .75 : displayHeight * image.width / image.height);
          const height = displayHeight ?? width * image.height / image.width;
          if (!Number.isFinite(width) || width <= 0 || !Number.isFinite(height) || height <= 0)
            throw new RangeError("Image display dimensions must be finite and positive.");
          await flushParagraphs(pages.graphicHeight(width, height));
          pages.registerXObject(image.name, image.id);
          await pages.graphic(width, height, alignment, (x, y, w, h) => "q " + [w, 0, 0, h, x, y].map(n).join(" ") + " cm /" + image.name + " Do Q\n", alternative); break;
        }
        case "vector": {
          const alternative = graphicAlternative(block, prepared), alignment = graphicAlignment(block.alignment), path = await vectorCommands(block, prepared, commands); commands += path.count;
          await flushParagraphs(pages.graphicHeight(path.width, path.height));
          await pages.graphic(path.width, path.height, alignment, (x, y, w, h) => "q " + [w / path.width, 0, 0, -h / path.height, x, y + h].map(n).join(" ") +
            " cm 0 0 " + n(path.width) + " " + n(path.height) + " re W n\n" + path.content + "Q\n", alternative); break;
        }
        case "drawing": {
          checkDrawingBlock(block);
          const alternative = graphicAlternative(block, prepared), alignment = graphicAlignment(block.alignment);
          const drawing = await prepareDrawing(block.drawing, prepared, fonts, { commands, elements: drawingElements });
          commands += drawing.commands; drawingElements += drawing.elements;
          await flushParagraphs(pages.graphicHeight(drawing.width, drawing.height));
          await pages.graphic(drawing.width, drawing.height, alignment, (x, y, w, h) => "q " + [w / drawing.width, 0, 0, -h / drawing.height, x, y + h].map(n).join(" ") +
            " cm 0 0 " + n(drawing.width) + " " + n(drawing.height) + " re W n\n" + drawing.content + "Q\n", alternative); break;
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
