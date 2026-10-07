// Product assertions; installation, sessions and captures belong to HtmlTinkerX.
async function emitFixture(name, blob) {
  let bytes;
  try { bytes = new Uint8Array(await blob.arrayBuffer()); }
  catch (error) { throw new Error("Reading fixture " + name + " failed: " + error); }
  let binary = "";
  for (let i = 0; i < bytes.length; i += 16384) binary += String.fromCharCode(...bytes.subarray(i, i + 16384));
  await writeFixture(name, btoa(binary));
}

async function runBrowserScenarios({ vectorJson, workerScript, limits }) {
  const vectors = JSON.parse(vectorJson);
  const { Workbook, writeCsv } = OfficeIMO;
  let assertions = 0;
  function require(value, message) { assertions++; if (!value) throw new Error(message); }
  async function rejects(action, kind) {
    try { await action(); } catch (error) { require(error.name === kind, "Expected " + kind + ", got " + error); return; }
    throw new Error("Expected rejection: " + kind);
  }
  const columns = [
    { header: "Name", key: "name", width: 28, wrapText: true, alignment: "left" },
    { header: "Site", key: "site", width: 18 },
    { header: "Last seen", key: "seen", type: "date", format: "yyyy-mm-dd hh:mm" },
    { header: "Latency (ms)", key: "latency", type: "number", format: "0.00", alignment: "right" },
    { header: "Healthy", key: "healthy", type: "boolean" }
  ];
  const expectedNames = ["Domain controllers", "domain controllers (2)", "Bad_______", "Sheet", "History_", "a".repeat(30)];
  for (const compression of ["auto", "store"]) {
    const events = [];
    const book = new Workbook({ creator: "Test<&\" 🧪", title: "Report <>&", created: new Date("2026-10-05T10:00:00Z"),
      modified: new Date("2026-10-05T11:00:00Z"), dateMode: "utc", compression, onProgress: p => events.push(p) });
    const sheet = book.addWorksheet(expectedNames[0], { columns, autoFilter: true, freezeHeader: true, headerFill: "D9E1F2" });
    await sheet.addRows([["DC<&\"'\r\n\t🧪שלום\u0001\ud800", "Łódź", new Date("2026-10-05T12:34:56Z"), 12.5, true]]);
    async function* additional() {
      yield { name: "=literal", site: "مرحبا", seen: new Date("1900-02-28T12:00:00Z"), latency: -1.25, healthy: false };
      yield { name: "_x0041_", site: null, seen: new Date("1900-03-01T00:00:00Z"), latency: Infinity, healthy: null };
      yield { name: "last", site: undefined, seen: new Date("1900-01-01T00:00:00Z"), latency: NaN, healthy: true };
    }
    await sheet.addRows(additional());
    const requestedNames = ["domain controllers", "'Bad[]:*?/\\'", "'  '", "History", "a".repeat(30) + "🧪"];
    requestedNames.forEach((requested, index) => require(book.addWorksheet(requested).name === expectedNames[index + 1], "Sheet name rules"));
    const blob = await book.toBlob();
    require(events.at(-1).phase === "complete" && events.at(-1).rows === 4, "Final progress count");
    require(await book.toBlob() === blob, "Idempotent finalization");
    await emitFixture("rich-" + compression + ".xlsx", blob);
  }
  await emitFixture("empty.xlsx", await new Workbook().toBlob());
  const one = new Workbook(), only = one.addWorksheet("One", { columns: [{ header: "V" }], includeHeader: false });
  await only.addRows([["one"]]); await emitFixture("one.xlsx", await one.toBlob());
  const long = new Workbook(), longSheet = long.addWorksheet("Long", { columns: [{ header: "Value" }] });
  await longSheet.addRows([["a".repeat(32765) + "🧪"]]); await emitFixture("long.xlsx", await long.toBlob());
  const wide = new Workbook(), wideSheet = wide.addWorksheet("Wide", { columns: Array.from({ length: 16384 }, () => ({ header: "V" })), includeHeader: false });
  const wideRow = Array(16384).fill(null); wideRow[16383] = "last";
  await wideSheet.addRows([wideRow]); await emitFixture("wide.xlsx", await wide.toBlob());
  for (const dateMode of ["local", "utc"]) {
    const book = new Workbook({ dateMode }), sheet = book.addWorksheet("Date", { columns: [{ header: "Date", type: "date" }], includeHeader: false });
    await sheet.addRows([[new Date("2026-10-05T12:34:56.123Z")]]);
    await emitFixture("date-" + dateMode + ".xlsx", await book.toBlob());
  }
  const original = globalThis.CompressionStream;
  try {
    for (const [name, constructor] of [["fallback", undefined], ["fallback-raw", class { constructor() { throw new TypeError("No raw deflate support"); } }]]) {
      globalThis.CompressionStream = constructor;
      const book = new Workbook(), sheet = book.addWorksheet("Fallback", { columns: [{ header: "Text" }] });
      await sheet.addRows([["fallback"]]); await emitFixture(name + ".xlsx", await book.toBlob());
    }
  } finally { globalThis.CompressionStream = original; }
  await rejects(async () => new Workbook().addWorksheet("Wide", { columns: Array(16385).fill({ header: "V" }) }), "RangeError");
  await rejects(async () => new Workbook().addWorksheet("Long", { columns: [{ header: "V" }] }).addRows([["a".repeat(32768)]]), "RangeError");
  const nativeChannel = globalThis.MessageChannel, nativeScheduler = globalThis.scheduler;
  const schedulerDescriptor = Object.getOwnPropertyDescriptor(globalThis, "scheduler");
  try {
    for (const scheduling of ["native", "channel", "timer"]) {
      Object.defineProperty(globalThis, "scheduler", { configurable: true, value: scheduling === "native" ? nativeScheduler : undefined });
      globalThis.MessageChannel = scheduling === "timer" ? undefined : nativeChannel;
      for (const kind of ["xlsx", "csv"]) {
        const controller = new AbortController(); let returned = false;
        function* rows() { try { for (let i = 0; i < 1000000; i++) yield ["row" + i]; } finally { returned = true; } }
        const options = { columns: [{ header: "V" }], signal: controller.signal };
        const book = new Workbook(options);
        const writing = kind === "xlsx" ? book.addWorksheet("Cancelled", options).addRows(rows()) : writeCsv(rows(), options);
        setTimeout(() => controller.abort(), 15);
        await rejects(() => writing, "AbortError"); require(returned, scheduling + " cancelled iterator was returned");
      }
    }
  } finally {
    globalThis.MessageChannel = nativeChannel;
    if (schedulerDescriptor) Object.defineProperty(globalThis, "scheduler", schedulerDescriptor);
    else delete globalThis.scheduler;
  }
  if (limits) {
    const book = new Workbook(), sheet = book.addWorksheet("Limit", { columns: [{ header: "V" }] });
    let consumed = 0;
    function* rows() { for (let i = 0; i < 1048576; i++) { consumed++; yield [i]; } }
    await rejects(() => sheet.addRows(rows()), "RangeError");
    require(consumed === 1048576, "Row limit includes header");
    await rejects(() => book.toBlob(), "RangeError");
  }
  for (const vector of vectors.cases) {
    const rows = vector.rows.map(row => row.map(v => v?.kind === "date" ? new Date(v.value) : v));
    await emitFixture(vector.name + ".csv", await writeCsv(rows, vector));
  }
  // The normal classic script also works in a host-owned Blob worker,
  // without introducing a worker-specific product API or another shipped runtime.
  const workerUrl = URL.createObjectURL(new Blob([workerScript, `
    const { Workbook, writeCsv } = OfficeIMO;
    globalThis.onmessage = async ({ data: rows }) => {
      try {
        const columns = [{ header: "Name" }, { header: "Date", type: "date" }, { header: "Value" }, { header: "Healthy" }];
        async function workbook(data) {
          const book = new Workbook({ dateMode: "utc" });
          await book.addWorksheet("Worker", { columns }).addRows(rows);
          book.addPart({ uri: "/customXml/worker.xml", contentType: "application/xml", data,
            relationship: { id: "workerData", type: OfficeIMO.opc.relationshipTypes.customXml } });
          return book.toBlob();
        }
        const xml = '<data xmlns="urn:worker">Łódź</data>';
        let xlsx, blobReadError;
        try { xlsx = await workbook(new Blob([xml])); }
        catch (error) {
          if (error.code !== "PLATFORM_UNAVAILABLE" || error.cause?.name !== "NotReadableError") throw error;
          blobReadError = { code: error.code, cause: error.cause.name, message: error.message };
          xlsx = await workbook(new TextEncoder().encode(xml));
        }
        postMessage({ xlsx, blobReadError, csv: await writeCsv(rows, { columns }) });
      } catch (error) { postMessage({ error: String(error) }); }
    };`], { type: "text/javascript" }));
  const worker = new Worker(workerUrl);
  let timeout;
  try {
    const data = await new Promise((resolve, reject) => {
      timeout = setTimeout(() => reject(new Error("Worker export timed out")), 10000);
      worker.onmessage = event => event.data.error ? reject(new Error(event.data.error)) : resolve(event.data);
      worker.onerror = event => reject(new Error("Worker error: " + event.message + " at " + event.filename + ":" + event.lineno));
      worker.postMessage([["Łódź", new Date("2026-10-05T12:34:56Z"), -2, true]]);
    });
    require(await data.csv.text() === "Name,Date,Value,Healthy\r\nŁódź,2026-10-05T12:34:56.000Z,-2,True\r\n", "Worker CSV values differ");
    if (data.blobReadError) require(data.blobReadError.code === "PLATFORM_UNAVAILABLE" && /byte chunks/.test(data.blobReadError.message), "Blocked worker Blob input did not report the host limitation");
    globalThis.workerBlobReadError = data.blobReadError ?? null;
    await emitFixture("worker.xlsx", data.xlsx);
  } finally { clearTimeout(timeout); worker.terminate(); URL.revokeObjectURL(workerUrl); }
  return { assertions, deflateRaw: (() => { try { return !!new CompressionStream("deflate-raw"); } catch { return false; } })(),
    rowLimitChecked: !!limits, workerChecked: true, workerBlobReadError: globalThis.workerBlobReadError,
    timezone: Intl.DateTimeFormat().resolvedOptions().timeZone };
}

async function runScale(format) {
  const columns = Array.from({ length: 20 }, (_, i) => ({ header: "Column " + i, type: i % 3 === 0 ? "number" : "string" }));
  function* rows() { for (let r = 0; r < 100000; r++) yield columns.map((_, c) => c % 3 === 0 ? r * 20 + c : c % 3 === 1 ? "Site " + r % 20 : "Row " + r + " column " + c); }
  const gaps = [], tasks = []; let last = performance.now();
  const timer = setInterval(() => { const now = performance.now(); gaps.push(now - last); last = now; }, 10);
  let observer;
  if (PerformanceObserver.supportedEntryTypes.includes("longtask")) {
    observer = new PerformanceObserver(list => tasks.push(...list.getEntries().map(e => e.duration)));
    observer.observe({ type: "longtask" });
  }
  const start = performance.now(), before = performance.memory?.usedJSHeapSize;
  let peak = before;
  const onProgress = () => { const heap = performance.memory?.usedJSHeapSize; if (heap !== undefined) peak = Math.max(peak ?? 0, heap); };
  let blob;
  try {
    if (format === "xlsx") {
      const book = new OfficeIMO.Workbook({ onProgress });
      await book.addWorksheet("Scale", { columns, autoFilter: true, freezeHeader: true }).addRows(rows());
      blob = await book.toBlob();
    } else blob = await OfficeIMO.writeCsv(rows(), { columns, onProgress });
    await new Promise(r => setTimeout(r, 20));
  } finally { clearInterval(timer); observer?.disconnect(); }
  globalThis.scaleMetrics = { format, rows: 100000, columns: 20, elapsedMs: performance.now() - start, bytes: blob.size,
    timerSamples: gaps.length, maxTimerGapMs: Math.max(...gaps), longTasksSupported: !!observer,
    longTasks: observer ? tasks.length : null, maxLongTaskMs: observer ? Math.max(0, ...tasks) : null,
    heapBefore: before ?? null, heapPeakSample: peak ?? null, heapAfter: performance.memory?.usedJSHeapSize ?? null };
  OfficeIMO.saveBlob(blob, "scale." + format);
}
