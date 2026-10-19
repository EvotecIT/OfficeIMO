import { BlobByteSink, withDestination } from "../core/sinks.js";
import { beginTask, checkAbort, consumeRows, pause, taskYieldDue } from "../core/iteration.js";
import { notifyProgress } from "../core/progress.js";
import type { ExportResult, ExportValue, OutputDestination } from "../core/index.js";
import { copyColumns, createRowProjector } from "../internal/rows.js";
import { settings, presentation, positive } from "./settings.js";
import { pdfNumber as n } from "./objects.js";
import { PdfTableLayout } from "./table-layout.js";
import { PdfWriteSession } from "./session.js";
import { PdfDocumentImages } from "./document-images.js";
import { vectorCommands } from "./document-vectors.js";
import type { PdfDocumentBlock, PdfDocumentOptions, PdfTableBlock, PdfTableOptions } from "./document-types.js";

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
    let rows = 0, columns = 0, commands = 0;
    await session.start();
    await consumeRows(blocks, options.signal, async block => {
      if (!block || typeof block !== "object") throw new TypeError("PDF document blocks must be objects.");
      switch (block.kind) {
        case "paragraph": {
          if (typeof block.text !== "string") throw new TypeError("Paragraph text must be a string.");
          if (block.presentation && Object.keys(block.presentation).some(key => !["bold", "italic", "color", "alignment"].includes(key)))
            throw new TypeError("PDF paragraph presentation supports bold, italic, color and alignment.");
          const style = presentation(block.presentation), size = positive(block.fontSize ?? prepared.fontSize, "Paragraph font size", 144);
          if (block.text) await pages.paragraph(block.text, style, size);
          break;
        }
        case "pageBreak": await pages.pageBreak(); break;
        case "table": {
          checkTableOptions(block.options);
          const table = settings({ ...options, ...block.options, columns: block.options?.columns, includeHeader: block.options?.includeHeader ?? true }), layout = new PdfTableLayout({ ...table, budget }, fonts);
          const project = createRowProjector(table.options.columns, undefined, options.signal);
          let count = 0; columns = Math.max(columns, table.options.columns.length);
          await pages.beginTable(layout);
          await consumeRows(block.rows, options.signal, async row => {
            budget.row(rows + 1);
            await pages.row(await layout.data(project(row, count) as readonly ExportValue[], count)); count++; rows++;
            if (rows % 256 === 0) notifyProgress(options.onProgress, { phase: "rows", rows, bytes: objects.bytes });
            if (taskYieldDue()) await pause(); checkAbort(options.signal);
          });
          const footer = await layout.footer(count);
          if (footer) { if (table.options.footer?.rows) await pages.block(footer); else await pages.row(footer[0]!); }
          pages.endTable(); break;
        }
        case "image": {
          const alignment = graphicAlignment(block.alignment), image = await images.prepare(block.image);
          const width = positive(block.width ?? (block.height === undefined ? image.width * .75 : block.height * image.width / image.height), "Image display width");
          const height = positive(block.height ?? width * image.height / image.width, "Image display height");
          pages.registerXObject(image.name, image.id);
          await pages.graphic(width, height, alignment, (x, y, w, h) => "q " + [w, 0, 0, h, x, y].map(n).join(" ") + " cm /" + image.name + " Do Q\n"); break;
        }
        case "vector": {
          const alignment = graphicAlignment(block.alignment), path = await vectorCommands(block, prepared, commands); commands += block.commands.length;
          await pages.graphic(block.width, block.height, alignment, (x, y, w, h) => "q " + [w / block.width, 0, 0, -h / block.height, x, y + h].map(n).join(" ") +
            " cm 0 0 " + n(block.width) + " " + n(block.height) + " re W n\n" + path + "Q\n"); break;
        }
        default: throw new TypeError("Unsupported PDF document block.");
      }
      if (taskYieldDue()) await pause(); checkAbort(options.signal);
    });
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
  const allowed = ["columns", "includeHeader", "headerRows", "footer", "headerPresentation", "footerPresentation", "alternateRowColor", "columnWidths", "wideTable", "formatValue"];
  if (!options || typeof options !== "object" || Object.keys(options).some(key => !allowed.includes(key)))
    throw new TypeError("Table options support columns, headings, footers, widths and presentation; document settings belong on writePdfDocument.");
}
