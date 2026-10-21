import { checkAbort, withAbort } from "../core/iteration.js";
import { OfficeIMOError } from "../core/errors.js";
import type { ExportProgress, ExportValue } from "../core/index.js";
import { bounds, captureInput, captureProgress, message, protocol, restoreError, restoreRows, result, wireError } from "./protocol.js";
import type { WireMessage } from "./protocol.js";
import type { ExportWorkerEndpoint, WorkerExportHandler, WorkerExportOperation } from "./types.js";

const installed = new WeakSet<ExportWorkerEndpoint>();

/** Install reusable transport on a host-created worker. The operation owns format selection and ordinary writer options. */
export function handleWorkerExports(endpoint: ExportWorkerEndpoint, operation: WorkerExportOperation): WorkerExportHandler {
  if (installed.has(endpoint)) throw new OfficeIMOError("INVALID_STATE", "A worker export handler is already installed.");
  if (typeof operation !== "function") throw new TypeError("A worker export operation is required.");
  installed.add(endpoint);
  let current: { readonly job: string; readonly controller: AbortController; readonly deliver: (data: WireMessage) => void } | undefined;
  let disposed = false;
  const post = (job: string, payload: Omit<WireMessage, "protocol" | "job">) => endpoint.postMessage({ protocol, job, ...payload });
  async function run(start: WireMessage): Promise<void> {
    const controller = new AbortController(), limits = bounds(start.bounds ?? {}), input = captureInput(start.input, limits);
    const pending = new Map<number, { resolve: (value: WireMessage) => void; reject: (error: unknown) => void }>();
    let sequence = 0, writing = false, claimed = false, latest: ExportProgress | undefined, progress: Promise<void> | undefined;
    const cancel = () => { for (const entry of pending.values()) entry.reject(controller.signal.reason); pending.clear(); latest = undefined; };
    controller.signal.addEventListener("abort", cancel, { once: true });
    current = { job: start.job, controller, deliver(data) {
      const entry = pending.get(data.id!);
      if (!entry) return;
      pending.delete(data.id!); entry.resolve(data);
    } };
    const ask = (kind: string, payload: Omit<WireMessage, "protocol" | "job" | "kind" | "id"> = {}): Promise<WireMessage> => {
      checkAbort(controller.signal);
      return new Promise((resolve, reject) => {
        const id = ++sequence; pending.set(id, { resolve, reject });
        try { post(start.job, { kind, id, ...payload }); } catch (error) { pending.delete(id); reject(error); }
      });
    };
    function onProgress(value: ExportProgress): void {
      checkAbort(controller.signal); latest = captureProgress(value, limits);
      if (!progress) {
        progress = (async () => {
          while (latest) { const next = latest; latest = undefined; await ask("progress", { progress: next }); }
        })().finally(() => { progress = undefined; });
        void progress.catch(error => controller.abort(error));
      }
    }
    const rows: AsyncIterable<readonly ExportValue[]> = { [Symbol.asyncIterator]() {
      if (claimed) throw new TypeError("Worker rows can be consumed only once."); claimed = true;
      return (async function* () {
        while (true) {
          const batch = await ask("requestRows");
          if (batch.kind !== "rows" || typeof batch.done !== "boolean") throw new TypeError("Invalid worker row response.");
          for (const row of restoreRows(batch.rows, limits)) { checkAbort(controller.signal); yield row; }
          if (batch.done) return;
        }
      })();
    } };
    const sink = { async write(bytes: Uint8Array): Promise<void> {
      checkAbort(controller.signal);
      if (writing || !(bytes instanceof Uint8Array)) throw new OfficeIMOError("INVALID_STATE", "Await each worker output write.");
      writing = true;
      try {
        for (let offset = 0; offset < bytes.length; offset += limits.maxChunkBytes)
          // Structured clone copies a view's entire backing buffer. Own exactly
          // this chunk so a large producer write cannot multiply clone memory.
          await ask("output", { bytes: new Uint8Array(bytes.subarray(offset, Math.min(bytes.length, offset + limits.maxChunkBytes))) });
      } finally { writing = false; }
    } };
    try {
      const completed = result(await withAbort(operation({ input, rows, sink, signal: controller.signal, onProgress }), controller.signal));
      if (progress) await withAbort(progress, controller.signal);
      if (pending.size) throw new OfficeIMOError("INVALID_STATE", "Await worker row and output operations before completing.");
      post(start.job, { kind: "complete", result: completed });
    } catch (error) {
      controller.abort(error);
      try { post(start.job, { kind: "failed", error: wireError(error) }); } catch { /* A closed endpoint cannot receive a terminal message. */ }
    } finally { controller.signal.removeEventListener("abort", cancel); cancel(); current = undefined; }
  }
  const receive: EventListener = event => {
    const data = message((event as MessageEvent).data); if (!data || disposed) return;
    if (data.kind === "start") {
      if (current) { post(data.job, { kind: "failed", error: wireError(new OfficeIMOError("INVALID_STATE", "The worker already has an active export.")) }); return; }
      void run(data).catch(error => { try { post(data.job, { kind: "failed", error: wireError(error) }); } catch { /* Disconnected endpoint. */ } });
    } else if (data.job === current?.job) {
      if (data.kind === "cancel") current.controller.abort(restoreError(data.error));
      else current.deliver(data);
    }
  };
  const broken: EventListener = event => current?.controller.abort(new Error((event as ErrorEvent).message || "Worker messaging failed."));
  endpoint.addEventListener("message", receive); endpoint.addEventListener("error", broken); endpoint.addEventListener("messageerror", broken); endpoint.start?.();
  return { dispose(reason = new DOMException("Worker export handler disposed.", "AbortError")) {
    if (disposed) return; disposed = true; current?.controller.abort(reason);
    endpoint.removeEventListener("message", receive); endpoint.removeEventListener("error", broken); endpoint.removeEventListener("messageerror", broken); installed.delete(endpoint);
  } };
}
