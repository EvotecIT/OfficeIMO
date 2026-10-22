import { BlobByteSink, withDestination } from "../core/sinks.js";
import { beginTask, checkAbort, consumeRows, pause, taskYieldDue } from "../core/iteration.js";
import { notifyProgress } from "../core/progress.js";
import type { ExportResult, OutputDestination, ExportValue } from "../core/index.js";
import { createRowProjector } from "../internal/rows.js";
import { settings } from "./settings.js";
import { PdfWriteSession } from "./session.js";
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
    const session = new PdfWriteSession(sink, prepared, true), { objects, pages } = session, layout = session.layout!;
    const project = createRowProjector(options.columns, undefined, options.signal);
    let count = 0;
    await session.start();
    const completed = (): void | Promise<void> => {
      count++;
      if (count % 256 === 0) {
        notifyProgress(options.onProgress, { phase: "rows", rows: count, bytes: objects.bytes });
        if (taskYieldDue()) return pause().then(() => { checkAbort(options.signal); });
      }
      checkAbort(options.signal);
    };
    await consumeRows(rows, options.signal, async row => {
      budget.row(count + 1);
      const values = project(row, count) as readonly ExportValue[];
      return pages.row(await layout.data(values, count)).then(completed);
    });
    const footer = await layout.footer(count);
    if (footer) {
      if (options.footer?.rows) await pages.block(footer);
      else await pages.row(footer[0]!);
    }
    if (layout.bands) await pages.endTable();
    if (options.messageBottom) await pages.paragraph(options.messageBottom);
    await session.finish();
    notifyProgress(options.onProgress, { phase: "complete", rows: count, bytes: objects.bytes });
    checkAbort(options.signal);
    return { rows: count, columns: options.columns.length, bytes: objects.bytes };
  });
}
