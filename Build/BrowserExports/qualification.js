// Opt-in artifact qualification. Timing is diagnostic; no host-dependent pass/fail thresholds.
async function runQualificationCase(args) {
  if (args.worker) return runQualificationWorker(args);
  if (args.cancelAfterRows === null) delete args.cancelAfterRows;
  const nativeCompression = globalThis.CompressionStream;
  if (args.fallback) globalThis.CompressionStream = class { constructor() { throw new TypeError("Qualification: deflate-raw unavailable"); } };
  const controller = new AbortController(), failure = new Error("qualification cancellation");
  let produced = 0, returned = false, pages = 0, outputBytes = 0, writes = 0, firstByteRows = null, maxChunk = 0, maxGap = 0;
  let peakHeap = performance.memory?.usedJSHeapSize ?? null;
  let previous = performance.now();
  const timer = setInterval(() => { const now = performance.now(); maxGap = Math.max(maxGap, now - previous); previous = now; if (peakHeap !== null) peakHeap = Math.max(peakHeap, performance.memory.usedJSHeapSize); }, 10);
  let timeout;
  const started = performance.now();
  let buffer = new Uint8Array(65536), used = 0;
  async function flush() {
    if (!used) return;
    let binary = "";
    for (let i = 0; i < used; i += 8192) binary += String.fromCharCode(...buffer.subarray(i, Math.min(used, i + 8192)));
    await globalThis.acceptQualificationChunk(btoa(binary));
    if (args.slowSink) await new Promise(resolve => setTimeout(resolve, 1));
    used = 0;
  }
  const sink = { async write(bytes) {
    firstByteRows ??= produced; outputBytes += bytes.length; maxChunk = Math.max(maxChunk, bytes.length); writes++;
    if (args.hangSink) { timeout = setTimeout(() => controller.abort(failure), 50); await new Promise(() => {}); }
    let offset = 0;
    while (offset < bytes.length) {
      const count = Math.min(bytes.length - offset, buffer.length - used);
      buffer.set(bytes.subarray(offset, offset + count), used); used += count; offset += count;
      if (used === buffer.length) await flush();
    }
  } };
  const { ExportCell, createWorkbook, writeCsvTo } = OfficeIMO;
  const columns = Array.from({ length: args.columns }, (_, c) => ({ header: "Column " + c, key: "c" + c,
    type: ["number", "string", "boolean", "date"][c % 4], ...(c % 4 === 0 ? { format: "0.00" } : {}), ...(c % 4 === 3 ? { format: "yyyy-mm-dd" } : {}),
    ...(args.styled ? { groups: [c < args.columns / 2 ? "Identity" : "Metrics"] } : {}) }));
  async function* source() {
    try {
      for (let r = 0; r < args.rows; r++) {
        if (r % 10000 === 0) { pages++; await Promise.resolve(); }
        if (args.cancelAfterRows !== undefined && produced === args.cancelAfterRows) controller.abort(failure);
        const row = columns.map((_, c) => {
          let value = c % 4 === 0 ? r * args.columns + c : c % 4 === 1 ? args.unique ? "Unique " + r + ": Łódź🧪" : "Site " + r % 8 : c % 4 === 2 ? r % 2 === 0 : new Date(Date.UTC(2026, 0, 1 + r % 28));
          if (args.longText && r === 100 && c === 1) value = "a".repeat(32766) + "🧪" + "Łódź\r\nשלום_x0041_".repeat(2500);
          return args.styled && c % 4 === 0 ? new ExportCell(value, { text: String(value), presentation: r % 3 === 0 ? { background: "FFF2CC", bold: true } : { color: "1F4E78" } }) : value;
        });
        produced++; yield row;
      }
    } finally { returned = true; }
  }
  let result, rejected = null;
  try {
    const options = { signal: controller.signal, ...(args.resourceLimit ? { limits: { maxRows: 5 } } : {}) };
    if (args.format === "csv") await writeCsvTo(source(), sink, { ...options, columns });
    else {
      const book = createWorkbook({ ...options, sink, dateMode: "utc", oversizedText: "preserve" });
      const sheet = book.addSheet("Qualified", { columns, ...(args.styled ? { table: { name: "QualifiedData" }, freezeHeader: true,
        autoSize: { sampleRows: 100, minWidth: 8, maxWidth: 40 }, alternatingRowStyle: { fill: { color: "E2F0D9" } },
        footer: { values: ["Totals"], totals: Object.fromEntries(columns.flatMap((_, c) => c % 4 === 0 ? [["c" + c, "sum"]] : [])), style: { font: { bold: true } } },
        print: { repeatHeaders: true, paper: "A4", orientation: "landscape" } } : {}) });
      await sheet.addRows(source()); result = await book.finish();
    }
    await flush();
  } catch (error) {
    rejected = error.code ?? error.message;
    if (!(error === failure || args.resourceLimit && error.code === "RESOURCE_LIMIT")) throw error;
  } finally { clearInterval(timer); clearTimeout(timeout); globalThis.CompressionStream = nativeCompression; buffer = null; }
  if (produced && !returned) throw new Error("The row iterator was not returned.");
  if ((args.cancelAfterRows !== undefined || args.hangSink || args.resourceLimit) && !rejected) throw new Error("Expected failure did not occur.");
  if (!rejected && produced !== args.rows) throw new Error("Source row count differs.");
  if (!rejected && args.rows >= 10000 && firstByteRows >= args.rows) throw new Error("Output was buffered until source completion.");
  return { ...args, workerScript: undefined, qualificationScript: undefined, result, produced, returned, pages, outputBytes, writes, maxChunk,
    firstByteRows, elapsedMs: performance.now() - started, peakHeapBytes: peakHeap, maxTimerGapMs: maxGap, rejected };
}
async function runQualificationWorker(args) {
  const bootstrap = `const pending = new Map(); let sequence = 0;
globalThis.acceptQualificationChunk = chunk => new Promise((resolve, reject) => { const id = ++sequence; pending.set(id, {resolve,reject}); postMessage({id,chunk}); });
onmessage = async event => { if (event.data.ack) { const p = pending.get(event.data.ack); pending.delete(event.data.ack); event.data.error ? p.reject(new Error(event.data.error)) : p.resolve(); return; }
try { postMessage({result:await runQualificationCase(event.data)}); } catch (error) { postMessage({error:error.stack || String(error)}); } };`;
  const url = URL.createObjectURL(new Blob([args.workerScript, "\n", args.qualificationScript, "\n", bootstrap], { type: "text/javascript" }));
  const worker = new Worker(url);
  try { return await new Promise((resolve, reject) => {
    worker.onerror = event => reject(new Error(event.message));
    worker.onmessage = async event => {
      const message = event.data;
      if (message.chunk !== undefined) { try { await globalThis.acceptQualificationChunk(message.chunk); worker.postMessage({ack:message.id}); } catch (error) { worker.postMessage({ack:message.id,error:String(error)}); } }
      else if (message.error) reject(new Error(message.error)); else resolve({...message.result,worker:true});
    };
    worker.postMessage({...args,worker:false,workerScript:undefined,qualificationScript:undefined});
  }); } finally { worker.terminate(); URL.revokeObjectURL(url); }
}
