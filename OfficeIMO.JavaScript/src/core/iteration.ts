export function checkAbort(signal?: AbortSignal): void {
  if (signal?.aborted) throw signal.reason ?? new DOMException("Export cancelled.", "AbortError");
}

/** Race a pending producer or sink operation without leaving abort listeners behind. */
export function withAbort<T>(promise: PromiseLike<T>, signal?: AbortSignal): Promise<T> {
  if (!signal) return Promise.resolve(promise);
  return new Promise<T>((resolve, reject) => {
    const cleanup = () => signal.removeEventListener("abort", abort);
    const abort = () => { cleanup(); reject(signal.reason ?? new DOMException("Export cancelled.", "AbortError")); };
    signal.addEventListener("abort", abort, { once: true });
    Promise.resolve(promise).then(value => { cleanup(); resolve(value); }, error => { cleanup(); reject(error); });
    if (signal.aborted) abort();
  });
}

/** Consume once; return the iterator on failure or cancellation. Pass the signal to I/O producers too. */
export async function* inputRows<T>(input: Iterable<T> | AsyncIterable<T>, signal?: AbortSignal): AsyncGenerator<T> {
  checkAbort(signal);
  const iterator = (input as AsyncIterable<T>)?.[Symbol.asyncIterator]?.() ?? (input as Iterable<T>)?.[Symbol.iterator]?.();
  if (!iterator) throw new TypeError("Rows must be a synchronous or asynchronous iterable.");
  let done = false, failed = false;
  try {
    while (true) {
      checkAbort(signal);
      const item = await withAbort(Promise.resolve(iterator.next()), signal);
      checkAbort(signal);
      if (item.done) { done = true; return; }
      yield item.value;
    }
  } catch (error) { failed = true; throw error; }
  finally {
    if (!done && iterator.return) {
      try {
        const returned = iterator.return();
        if (signal?.aborted || failed) Promise.resolve(returned).catch(() => {});
        else await withAbort(Promise.resolve(returned), signal);
      } catch (error) { if (!signal?.aborted && !failed) throw error; }
    }
  }
}

type RowConsumer = (signal: AbortSignal | undefined, accept: (value: unknown) => void | Promise<void>) => Promise<void>;
const rowConsumers = new WeakMap<object, RowConsumer>();

/** @internal Keep bounded pages as ordinary async rows, with direct batch consumption for owned writers. */
export function rowsFromBatches<T>(batches: Iterable<readonly T[]> | AsyncIterable<readonly T[]>): AsyncIterable<T> {
  let consumed = false;
  const claim = () => { if (consumed) throw new TypeError("A row source can be consumed only once."); consumed = true; };
  const rows: AsyncIterable<T> = { [Symbol.asyncIterator]() { claim(); return iterate(); } };
  async function* iterate(): AsyncGenerator<T> { for await (const batch of inputRows(batches)) yield* batch; }
  rowConsumers.set(rows, (signal, accept) => {
    claim(); return consumeRows(batches, signal, batch => consumeRows(batch, signal, accept));
  });
  return rows;
}

/** @internal Concatenate headings, body and footer without another per-row async delegation. */
export function concatRows<T>(...sources: (Iterable<T> | AsyncIterable<T>)[]): AsyncIterable<T> {
  const rows: AsyncIterable<T> = { async *[Symbol.asyncIterator]() { for (const source of sources) yield* source; } };
  rowConsumers.set(rows, async (signal, accept) => { for (const source of sources) await consumeRows(source, signal, accept); });
  return rows;
}

/** @internal Consume synchronous work without an async-generator and per-row Promise.
 * Async producers, promised values and destination backpressure keep the same cancellation/return contract. */
export async function consumeRows<T>(input: Iterable<T> | AsyncIterable<T>, signal: AbortSignal | undefined,
  accept: (value: T) => void | Promise<void>): Promise<void> {
  checkAbort(signal);
  const consume = rowConsumers.get(input);
  if (consume) return consume(signal, accept as (value: unknown) => void | Promise<void>);
  const iterator = (input as AsyncIterable<T>)?.[Symbol.asyncIterator]?.() ?? (input as Iterable<T>)?.[Symbol.iterator]?.();
  if (!iterator) throw new TypeError("Rows must be a synchronous or asynchronous iterable.");
  let done = false;
  try {
    while (true) {
      checkAbort(signal);
      let item: IteratorResult<T> | PromiseLike<IteratorResult<T>> = iterator.next();
      const next = item as PromiseLike<IteratorResult<T>>;
      if (typeof next?.then === "function") item = await withAbort(next, signal);
      checkAbort(signal);
      const result = item as IteratorResult<T>;
      if (result.done) { done = true; return; }
      let value = result.value;
      if (typeof (value as PromiseLike<T> | undefined)?.then === "function") value = await withAbort(value as PromiseLike<T>, signal);
      checkAbort(signal);
      const pending = accept(value);
      if (pending !== undefined) await withAbort(pending, signal);
    }
  } finally {
    if (!done && iterator.return) {
      // Any exit before exhaustion is a producer/consumer failure or cancellation.
      // Observe cleanup, but an unresponsive return must not replace or hold the original failure.
      try { void Promise.resolve(iterator.return()).catch(() => {}); } catch { /* Preserve the original failure. */ }
    }
  }
}

let taskDeadline: number | undefined;
const taskBudgetMs = 32;

/** @internal Start a new write phase without carrying an idle operation's expired deadline. */
export function beginTask(): void { taskDeadline = performance.now() + taskBudgetMs; }

/** @internal Pipeline stages share the last completed yield instead of pausing back-to-back. */
export function taskYieldDue(): boolean {
  const now = performance.now();
  taskDeadline ??= now + taskBudgetMs;
  return now >= taskDeadline;
}

/** Yield a task so input, rendering and cancellation can run without nested timer delays. */
export function pause(): Promise<void> {
  const scheduler = (globalThis as { scheduler?: { postTask?: (callback: () => void, options: { priority: "background" }) => Promise<void> } }).scheduler;
  // Boosted yield continuations can starve cancellation timers. A background task
  // lets due timers and input run before the next bounded section of export work.
  if (typeof scheduler?.postTask === "function") return scheduler.postTask(() => {
    taskDeadline = performance.now() + taskBudgetMs;
  }, { priority: "background" });
  return new Promise(resolve => {
    let done = false, channel: MessageChannel | undefined;
    const finish = () => {
      if (done) return;
      done = true; clearTimeout(timer);
      channel?.port1.close(); channel?.port2.close();
      taskDeadline = performance.now() + taskBudgetMs;
      resolve();
    };
    const timer = setTimeout(finish, 0);
    if (typeof MessageChannel === "function") {
      channel = new MessageChannel();
      channel.port1.onmessage = finish;
      channel.port2.postMessage(undefined);
    }
  });
}
