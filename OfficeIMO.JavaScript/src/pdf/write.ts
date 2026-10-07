import { BlobByteSink, withDestination } from "../core/sinks.js";
import { beginTask, checkAbort, consumeRows, pause, taskYieldDue } from "../core/iteration.js";
import type { ExportResult, OutputDestination, ExportValue } from "../core/index.js";
import { createRowProjector } from "../internal/rows.js";
import { settings } from "./settings.js";
import { PdfObjects, unicodeHex } from "./objects.js";
import { PdfFontResources } from "./font-resources.js";
import { PdfTableLayout } from "./table-layout.js";
import { PdfPages } from "./pages.js";
import type { PdfOptions } from "./types.js";

/** One table as a PDF Blob. Async sources are consumed once; call a source factory for each export. */
export function writePdf<T extends object>(rows: Iterable<T> | AsyncIterable<T>, options: PdfOptions<NoInfer<T>>): Promise<Blob>;
export async function writePdf(rows: Iterable<unknown> | AsyncIterable<unknown>, configuration: unknown): Promise<Blob> {
  const sink = new BlobByteSink();
  try { await writePdfTo(rows as Iterable<never>, sink, configuration as PdfOptions); return sink.toBlob("application/pdf"); }
  catch (error) { sink.discard(); throw error; }
}
/** Stream complete pages with backpressure. The destination remains caller-owned, including partial bytes on failure. */
export function writePdfTo<T extends object>(rows: Iterable<T> | AsyncIterable<T>, destination: OutputDestination, options: PdfOptions<NoInfer<T>>): Promise<ExportResult>;
export async function writePdfTo(rows: Iterable<unknown> | AsyncIterable<unknown>, destination: OutputDestination, configuration: unknown): Promise<ExportResult> {
  const prepared = settings(configuration as PdfOptions), { options, budget } = prepared;
  beginTask();
  checkAbort(options.signal);
  return withDestination(destination, async sink => {
    const objects = new PdfObjects(sink, options.signal, options.limits?.maxOutputBytes), root = objects.reserve(), parent = objects.reserve(), resources = objects.reserve(), info = objects.reserve();
    const fonts = new PdfFontResources(objects, options.fonts, prepared.limits.maxFontBytes), layout = new PdfTableLayout(prepared, fonts), pages = new PdfPages(objects, prepared, layout, parent, resources);
    const project = createRowProjector(options.columns, undefined, options.signal);
    let count = 0;
    await objects.text("%PDF-1.7\n");
    await objects.raw(Uint8Array.of(0x25, 0xe2, 0xe3, 0xcf, 0xd3, 0x0a));
    await pages.start();
    const completed = (): void | Promise<void> => {
      count++;
      if (count % 256 === 0) {
        options.onProgress?.({ phase: "rows", rows: count, bytes: objects.bytes });
        if (taskYieldDue()) return pause().then(() => { checkAbort(options.signal); });
      }
      checkAbort(options.signal);
    };
    await consumeRows(rows, options.signal, row => {
      budget.row(count + 1);
      const values = project(row, count) as readonly ExportValue[];
      return pages.row(layout.data(values, count)).then(completed);
    });
    const footer = layout.footer(count);
    if (footer) {
      if (options.footer?.rows) await pages.block(footer);
      else await pages.row(footer[0]!);
    }
    if (options.messageBottom !== undefined) await pages.paragraph(options.messageBottom);
    await pages.finish();
    for (const font of fonts.resources) await font.write(objects, options.compression !== false);
    await objects.object(resources, "<< /Font " + fonts.dictionary() + pages.xobjects() + " >>");
    await objects.object(parent, "<< /Type /Pages /Count " + pages.references.length + " /Kids [" + pages.references.map(id => id + " 0 R").join(" ") + "] >>");
    await objects.object(root, "<< /Type /Catalog /Pages " + parent + " 0 R >>");
    await objects.object(info, "<< /Producer (OfficeIMO.js)" + (options.title === undefined ? "" : " /Title <feff" + unicodeHex(options.title) + ">") + " >>");
    await objects.finish(root, info);
    options.onProgress?.({ phase: "complete", rows: count, bytes: objects.bytes });
    checkAbort(options.signal);
    return { rows: count, columns: options.columns.length, bytes: objects.bytes };
  });
}
