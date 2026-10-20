import { checkAbort, inputRows, withAbort } from "../core/iteration.js";
import { withDestination } from "../core/sinks.js";
import { notifyProgress } from "../core/progress.js";
import { OfficeIMOError } from "../core/errors.js";
import type { ExportResult, OutputDestination } from "../core/index.js";
import { bounds, captureInput, captureProgress, captureRow, message, protocol, restoreError, result, wireError } from "./protocol.js";
import type { WireCell, WireMessage } from "./protocol.js";
import type { ExportWorkerEndpoint, WorkerExportOptions, WorkerExportRows } from "./types.js";
import { captureFields } from "../internal/capture.js";

const active = new WeakSet<ExportWorkerEndpoint>();
let sequence = 0;

/** Execute through a host-owned worker with bounded rows, acknowledged bytes and coalesced progress.
 * The host owns the worker, source I/O and partial destination bytes. No caller-owned buffer is transferred. */
export async function executeWorkerExport(endpoint: ExportWorkerEndpoint, rows: WorkerExportRows, destination: OutputDestination,
  options: WorkerExportOptions = {}): Promise<ExportResult> {
  const captured = captureFields(options, ["input", "signal", "onProgress", "batchRows", "maxBatchCells", "maxBatchCharacters", "maxChunkBytes", "maxSetupBytes", "maxSetupNodes"]);
  const signal = captured.signal, onProgress = captured.onProgress;
  checkAbort(signal);
  const limits = bounds(captured), input = captureInput(captured.input, limits);
  if (active.has(endpoint)) throw new OfficeIMOError("INVALID_STATE", "A worker endpoint already has an active export.");
  active.add(endpoint);
  try { return await withDestination(destination, sink => new Promise<ExportResult>((resolve, reject) => {
    const lifetime = new AbortController();
    const job = "export-" + ++sequence, iterator = inputRows(rows, lifetime.signal);
    let settled = false, reading = false, writing = false, exhausted = false, acceptedBytes = 0;
    let pending: ReturnType<typeof captureRow> | undefined;
    const post = (payload: Omit<WireMessage, "protocol" | "job">) => endpoint.postMessage({ protocol, job, ...payload });
    function cleanup(): void {
      endpoint.removeEventListener("message", receive); endpoint.removeEventListener("error", broken); endpoint.removeEventListener("messageerror", broken);
      signal?.removeEventListener("abort", aborted);
      // A producer's stuck or rejected cleanup must not replace or delay export completion.
      if (!exhausted) void iterator.return(undefined).catch(() => {});
      pending = undefined;
    }
    function fail(error: unknown): void {
      if (settled) return;
      settled = true;
      // A terminal worker/transport failure must also release a pending producer
      // when the caller did not supply or abort its own signal.
      lifetime.abort(error);
      try { post({ kind: "cancel", error: wireError(error) }); } catch { /* A disconnected worker cannot receive cancellation. */ }
      cleanup(); reject(error);
    }
    const aborted = () => { try { checkAbort(signal); } catch (error) { fail(error); } };
    const broken: EventListener = event => fail(new Error((event as ErrorEvent).message || "Worker messaging failed."));
    async function read(id: number): Promise<void> {
      if (reading) throw new OfficeIMOError("INVALID_STATE", "Only one worker row request may be pending.");
      reading = true;
      try {
        const batch: (readonly WireCell[])[] = []; let cells = 0, characters = 0;
        while (batch.length < limits.batchRows && !exhausted) {
          let next = pending; pending = undefined;
          if (!next) {
            const item = await withAbort(iterator.next(), lifetime.signal);
            if (settled) return;
            if (item.done) { exhausted = true; break; }
            if (!Array.isArray(item.value) || item.value.length > limits.maxBatchCells)
              throw new OfficeIMOError("RESOURCE_LIMIT", "A worker row exceeds its message cell budget.");
            next = captureRow(item.value);
          }
          if (next.cells.length > limits.maxBatchCells || next.characters > limits.maxBatchCharacters)
            throw new OfficeIMOError("RESOURCE_LIMIT", "A worker row exceeds its message budget.");
          if (batch.length && (cells + next.cells.length > limits.maxBatchCells || characters + next.characters > limits.maxBatchCharacters)) { pending = next; break; }
          batch.push(next.cells); cells += next.cells.length; characters += next.characters;
        }
        if (!settled) post({ kind: "rows", id, rows: batch, done: exhausted });
      } finally { reading = false; }
    }
    async function write(data: WireMessage): Promise<void> {
      if (writing || !(data.bytes instanceof Uint8Array) || data.bytes.length > limits.maxChunkBytes)
        throw new OfficeIMOError("INVALID_STATE", "Invalid or overlapping worker output chunk.");
      writing = true;
      try {
        await withAbort(Promise.resolve(sink.write(data.bytes)), lifetime.signal);
        if (!settled) { acceptedBytes += data.bytes.length; post({ kind: "ack", id: data.id! }); }
      }
      finally { writing = false; }
    }
    const receive: EventListener = event => {
      const data = message((event as MessageEvent).data);
      if (settled || !data || data.job !== job) return;
      try {
        if (data.kind === "complete") {
          const completed = result(data.result);
          if (writing || reading || completed.bytes !== acceptedBytes) throw new OfficeIMOError("INVALID_STATE", "Worker completion does not match acknowledged output.");
          settled = true; cleanup(); resolve(completed);
        }
        else if (data.kind === "failed") fail(restoreError(data.error));
        else {
          if (!Number.isSafeInteger(data.id) || data.id! < 1) throw new TypeError("Invalid worker request identifier.");
          if (data.kind === "requestRows") void read(data.id!).catch(fail);
          else if (data.kind === "output") void write(data).catch(fail);
          else if (data.kind === "progress") { if (!data.progress) throw new TypeError("Invalid worker progress."); notifyProgress(onProgress, captureProgress(data.progress, limits)); post({ kind: "ack", id: data.id! }); }
          else throw new TypeError("Unexpected worker export message.");
        }
      } catch (error) { fail(error); }
    };
    endpoint.addEventListener("message", receive); endpoint.addEventListener("error", broken); endpoint.addEventListener("messageerror", broken);
    signal?.addEventListener("abort", aborted, { once: true });
    try { endpoint.start?.(); checkAbort(signal); post({ kind: "start", input, bounds: limits }); }
    catch (error) { fail(error); }
  })); } finally { active.delete(endpoint); }
}
