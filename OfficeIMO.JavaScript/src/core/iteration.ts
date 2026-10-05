export function checkAbort(signal?: AbortSignal): void {
  if (signal?.aborted) throw signal.reason ?? new DOMException("Export cancelled.", "AbortError");
}

/** Race a pending producer or sink operation without leaving abort listeners behind. */
export function withAbort<T>(promise: PromiseLike<T>, signal?: AbortSignal): Promise<T> {
  if (!signal) return Promise.resolve(promise);
  checkAbort(signal);
  return new Promise<T>((resolve, reject) => {
    const abort = () => reject(signal.reason ?? new DOMException("Export cancelled.", "AbortError"));
    signal.addEventListener("abort", abort, { once: true });
    Promise.resolve(promise).then(resolve, reject).finally(() => signal.removeEventListener("abort", abort));
  });
}

/** Consume once; return the iterator on failure or cancellation. Pass the signal to I/O producers too. */
export async function* inputRows<T>(input: Iterable<T> | AsyncIterable<T>, signal?: AbortSignal): AsyncGenerator<T> {
  checkAbort(signal);
  const iterator = (input as AsyncIterable<T>)?.[Symbol.asyncIterator]?.() ?? (input as Iterable<T>)?.[Symbol.iterator]?.();
  if (!iterator) throw new TypeError("Rows must be a synchronous or asynchronous iterable.");
  let done = false;
  try {
    while (true) {
      checkAbort(signal);
      const item = await withAbort(Promise.resolve(iterator.next()), signal);
      checkAbort(signal);
      if (item.done) { done = true; return; }
      yield item.value;
    }
  } finally {
    if (!done && iterator.return) {
      const returned = iterator.return();
      if (signal?.aborted) Promise.resolve(returned).catch(() => {});
      else await returned;
    }
  }
}

export function pause(): Promise<void> { return new Promise(resolve => setTimeout(resolve, 0)); }
