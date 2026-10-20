import type { ByteSink, ExportProgress, ExportResult, ExportValue } from "../core/index.js";

/** A host-created Worker, MessagePort or worker-global messaging surface. Ownership stays with the host. */
export interface ExportWorkerEndpoint {
  postMessage(message: unknown): void;
  addEventListener(type: string, listener: EventListener): void;
  removeEventListener(type: string, listener: EventListener): void;
  start?(): void;
}

/** Portable, already projected table rows. Live grids and domain callbacks remain in their owning host. */
export type WorkerExportRows = Iterable<readonly ExportValue[]> | AsyncIterable<readonly ExportValue[]>;

/** Bounds the messages in flight, independently of a format writer's document limits. */
export interface WorkerExportOptions {
  /** Structured-cloneable operation configuration. Functions and live grids cannot cross a worker boundary. */
  readonly input?: unknown;
  readonly signal?: AbortSignal;
  readonly onProgress?: (progress: ExportProgress) => void;
  /** At most this many rows in one request; defaults to 64. */
  readonly batchRows?: number;
  /** Includes every selected cell in a batch; defaults to 65,536. One row must fit. */
  readonly maxBatchCells?: number;
  /** UTF-16 units of values, display text, links and presentation strings; defaults to 1,000,000. */
  readonly maxBatchCharacters?: number;
  /** At most this many output bytes await destination acknowledgement; defaults to 65,536. */
  readonly maxChunkBytes?: number;
  /** UTF-16 string storage plus binary bytes in portable operation configuration; defaults to 16 MiB. */
  readonly maxSetupBytes?: number;
  /** Configuration values, array entries and object fields; defaults to 100,000, with depth limited to 32. */
  readonly maxSetupNodes?: number;
}

/** One worker operation; call the normal format writer with these rows, sink, signal and progress callback. */
export interface WorkerExportContext {
  readonly input: unknown;
  readonly rows: AsyncIterable<readonly ExportValue[]>;
  readonly sink: ByteSink;
  readonly signal: AbortSignal;
  readonly onProgress: (progress: ExportProgress) => void;
}
export type WorkerExportOperation = (context: WorkerExportContext) => Promise<ExportResult>;

/** Removes only the library's listeners and cancels the current operation; it never terminates the host's worker. */
export interface WorkerExportHandler { dispose(reason?: unknown): void; }
