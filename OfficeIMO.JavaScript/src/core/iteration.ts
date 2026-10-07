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

/** Yield a task so input, rendering and cancellation can run without nested timer delays. */
export function pause(): Promise<void> {
  if (typeof MessageChannel === "function") return new Promise(resolve => {
    const channel = new MessageChannel();
    channel.port1.onmessage = () => { channel.port1.close(); channel.port2.close(); resolve(); };
    channel.port2.postMessage(undefined);
  });
  return new Promise(resolve => setTimeout(resolve, 0));
}
