// Generated: Build/assets.mjs. MIT.
const utf8 = new TextEncoder();

function checkAbort(signal) {
  if (signal?.aborted) throw signal.reason ?? new DOMException("Export cancelled.", "AbortError");
}

function withAbort(promise, signal) {
  if (!signal) return promise;
  checkAbort(signal);
  return new Promise((resolve, reject) => {
    const abort = () => reject(signal.reason ?? new DOMException("Export cancelled.", "AbortError"));
    signal.addEventListener("abort", abort, { once: true });
    Promise.resolve(promise).then(resolve, reject).finally(() => signal.removeEventListener("abort", abort));
  });
}

async function* inputRows(input, signal) {
  checkAbort(signal);
  const iterator = input?.[Symbol.asyncIterator]?.() ?? input?.[Symbol.iterator]?.();
  if (!iterator) throw new TypeError("Rows must be a synchronous or asynchronous iterable.");
  let done = false;
  try {
    while (true) {
      checkAbort(signal);
      const next = iterator.next();
      const item = next?.then ? await withAbort(next, signal) : next;
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

function rowValues(row, columns) {
  if (Array.isArray(row)) {
    if (row.length > columns.length) throw new RangeError("Row has more values than declared columns.");
    return row;
  }
  if (!row || typeof row !== "object" || row instanceof Date) throw new TypeError("A row must be an array or object.");
  return columns.map(c => row[c.key ?? c.header]);
}

function copyColumns(columns) {
  if (!Array.isArray(columns)) throw new TypeError("Declare the columns in export order.");
  return columns.map(c => {
    if (!c || typeof c.header !== "string" || (c.key !== undefined && typeof c.key !== "string"))
      throw new TypeError("Each column needs a string header and an optional string key.");
    return { ...c };
  });
}

function pause() { return new Promise(resolve => setTimeout(resolve, 0)); }

function textChunks(write, signal) {
  let text = "", deadline = performance.now() + 8;
  async function flush() {
    checkAbort(signal);
    if (text) { const bytes = utf8.encode(text); text = ""; await write(bytes); }
    if (performance.now() >= deadline) { await pause(); deadline = performance.now() + 8; }
    checkAbort(signal);
  }
  return {
    append(value) { text += value; return text.length >= 32768 || performance.now() >= deadline; },
    flush
  };
}

function saveBlob(blob, fileName) {
  if (!(blob instanceof Blob) || typeof fileName !== "string" || !fileName) throw new TypeError("A Blob and file name are required.");
  const url = URL.createObjectURL(blob);
  const link = document.createElement("a");
  link.href = url; link.download = fileName;
  document.body.append(link);
  try { link.click(); } finally {
    link.remove();
    setTimeout(() => URL.revokeObjectURL(url), 30000);
  }
}

// OfficeIMO.CSV formula rule.
function csvField(value, delimiter, protect) {
  let text;
  if (value == null) text = "";
  else if (value instanceof Date) text = Number.isFinite(value.getTime()) ? value.toISOString() : "";
  else if (typeof value === "boolean") text = value ? "True" : "False";
  else if (typeof value === "number") text = Number.isFinite(value) ? String(value) : "";
  else if (typeof value === "string") text = value;
  else throw new TypeError("CSV cells must be strings, numbers, booleans, Dates or null.");
  if (protect && typeof value === "string" && /^ *[=+\-@\t\r\n]/.test(text)) text = "'" + text;
  return text.includes(delimiter) || /["\r\n]/.test(text) ? '"' + text.replace(/"/g, '""') + '"' : text;
}

async function writeCsv(rows, options = {}) {
  const columns = copyColumns(options.columns);
  const delimiter = options.delimiter ?? ",", lineEnding = options.lineEnding ?? "\r\n";
  if (![",", ";", "\t"].includes(delimiter)) throw new RangeError("Delimiter must be comma, semicolon or tab.");
  if (!["\r\n", "\n", "\r"].includes(lineEnding)) throw new RangeError("Invalid line ending.");
  const signal = options.signal, protect = options.formulaInjectionProtection !== false;
  checkAbort(signal);
  const parts = options.bom ? [new Uint8Array([239, 187, 191])] : [];
  const buffer = textChunks(bytes => { parts.push(bytes); }, signal);
  let count = 0;
  async function record(values) {
    for (let i = 0; i < columns.length; i++) {
      if (buffer.append((i ? delimiter : "") + csvField(values[i], delimiter, protect))) await buffer.flush();
    }
    if (buffer.append(lineEnding)) { await buffer.flush(); options.onProgress?.({ phase: "rows", rows: count }); }
  }
  if (options.includeHeader !== false && columns.length) await record(columns.map(c => c.header));
  for await (const row of inputRows(rows, signal)) {
    await record(rowValues(row, columns));
    count++;
  }
  await buffer.flush();
  const blob = new Blob(parts, { type: "text/csv;charset=utf-8" });
  options.onProgress?.({ phase: "complete", rows: count, bytes: blob.size });
  checkAbort(signal);
  return blob;
}

export { writeCsv, saveBlob };
