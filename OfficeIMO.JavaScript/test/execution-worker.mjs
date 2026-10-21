import { workerData } from "node:worker_threads";
import { handleWorkerExports } from "../dist/execution/index.js";
import { writeCsvTo, writeXlsxTo, writePdfTo, PdfFont } from "../dist/index.js";

handleWorkerExports(workerData.port, async ({ input, rows, sink, signal, onProgress }) => {
  const options = { columns: input.columns, signal, onProgress, ...(input.fontBytes ? { fonts: { regular: await PdfFont.create(input.fontBytes, { signal }) } } : {}) };
  return { csv: writeCsvTo, xlsx: writeXlsxTo, pdf: writePdfTo }[input.format](rows, sink, options);
});
