// One complete report operation per process. PowerForge owns repetitions and ordering.
import { readFile, writeFile, mkdir, open } from "node:fs/promises";
import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import { pathToFileURL, fileURLToPath } from "node:url";
import os from "node:os";
import path from "node:path";

const args = process.argv.slice(2), command = args.shift();
const opts = Object.fromEntries(args.flatMap((arg, i) => i % 2 ? [] : [[arg.replace(/^--/, ""), args[i + 1]]]));
const required = name => { if (!opts[name]) throw Error("Missing --" + name); return opts[name]; };
const sha = value => createHash("sha256").update(value).digest("hex");
const fixtures = [
  { name: "inventory-csv-100000-sink", format: "csv", rows: 100000, delivery: "sink", styled: false },
  { name: "inventory-xlsx-10000-sink", format: "xlsx", rows: 10000, delivery: "sink", styled: false },
  { name: "inventory-xlsx-100000-sink", format: "xlsx", rows: 100000, delivery: "sink", styled: false },
  { name: "editable-analysis-xlsx-100000-sink", format: "xlsx", rows: 100000, delivery: "sink", kind: "editable-analysis", currentOnly: true },
  { name: "assessment-xlsx-10000-blob", format: "xlsx", rows: 10000, delivery: "blob", styled: true },
  { name: "inventory-pdf-10000-sink", format: "pdf", rows: 10000, delivery: "sink", styled: false },
  { name: "inventory-pdf-100000-sink", format: "pdf", rows: 100000, delivery: "sink", styled: false },
  { name: "assessment-pdf-10000-blob", format: "pdf", rows: 10000, delivery: "blob", styled: true },
  { name: "tagged-pdf-10000-sink", format: "pdf", rows: 10000, delivery: "sink", styled: true, tagged: true },
  { name: "tagged-pdf-100000-sink", format: "pdf", rows: 100000, delivery: "sink", styled: true, tagged: true }
];
const titles = ["ID", "Asset", "Owner", "Status", "Amount", "Seen"];
function row(index) {
  return ["N" + String(index).padStart(6, "0"), "node-" + String(index).padStart(6, "0"), "ops-" + index % 17,
    ["Pass", "Warning", "Critical"][index % 3], index + 0.5, new Date("2026-01-" + String(index % 28 + 1).padStart(2, "0") + "T00:00:00Z")];
}
function* rows(api, fixture) {
  for (let index = 0; index < fixture.rows; index++) {
    const values = row(index);
    if (fixture.styled) values[3] = new api.ExportCell(values[3], { presentation: { background: ["E2F0D9", "FFF2CC", "FCE4D6"][index % 3], bold: index % 3 === 2 } });
    yield values;
  }
}
function settings(fixture) {
  const columns = titles.map((header, index) => ({ header, width: [14, 20, 12, 15, 14, 18][index], ...(index === 4 ? { type: "number", format: "0.0" } : index === 5 ? { type: "date", format: "yyyy-mm-dd" } : {}) }));
  if (fixture.format === "pdf") return { columns, pageSize: "A3", orientation: "landscape", fontSize: 8, padding: 2,
    pageNumbers: false, compression: true, formatValue: value => value instanceof Date ? value.toISOString().slice(0, 10) : String(value ?? ""), ...(fixture.tagged ? { tagged: { language: "en-US" } } : {}),
    ...(fixture.styled ? { alternateRowColor: "F3F6F9", headerPresentation: { bold: true, background: "D9E2F3" } } : {}) };
  if (fixture.format === "xlsx") return { columns, dateMode: "utc", sheet: { name: "Inventory", autoSize: { sampleRows: 100 },
    ...(fixture.styled ? { table: { name: "Assessment", style: "TableStyleMedium2" } } : {}) } };
  return { columns };
}
async function manifest() {
  const value = JSON.parse(await readFile(required("manifest"), "utf8"));
  if (sha(await readFile(fileURLToPath(import.meta.url))) !== value.helperSha256) throw Error("Workload helper changed after manifest preparation.");
  if (sha(await readFile(path.join(path.dirname(fileURLToPath(import.meta.url)), "verify-report-workload.py"))) !== value.verifierSha256) throw Error("Independent verifier changed after preparation.");
  if (value.example && sha(await readFile(value.example.path)) !== value.example.sha256) throw Error("Canonical editable-analysis example changed after preparation.");
  return value;
}
async function run() {
  const input = await manifest(), fixture = input.cases.find(item => item.name === required("case")), engine = required("engine");
  if (!fixture || !input.engines[engine]) throw Error("Unknown workload or engine.");
  if (fixture.currentOnly && engine !== "current") throw Error("Editable analysis is current-only diagnostics; the snapshot baseline is not equivalent work.");
  const source = input.engines[engine];
  if (sha(await readFile(source.path)) !== source.sha256) throw Error("Measured runtime changed after preparation.");
  const api = await import(pathToFileURL(source.path).href), config = settings(fixture), output = path.resolve(required("output"));
  const example = fixture.kind === "editable-analysis" ? await import(pathToFileURL(input.example.path).href) : undefined;
  const file = await open(output, "w"), memoryBefore = process.memoryUsage(), loadBefore = os.loadavg(), cpuStart = process.cpuUsage();
  let outputBytes = 0, firstByteMs = null, last = performance.now(), maxTimerGapMs = 0, timerTicks = 0;
  let produced = 0, enumerations = 0, returned = 0, writes = 0, firstDeliveryProducedRows, acknowledgementOverlapWrites = 0, maxRowsAdvancedDuringAcknowledgement = 0;
  function* selectedRows() {
    enumerations++;
    try { for (const value of example.operationalRows(fixture.rows)) { produced++; yield value; } }
    finally { returned++; }
  }
  let maxSampledHeapBytes = memoryBefore.heapUsed, maxSampledExternalBytes = memoryBefore.external;
  const timer = setInterval(() => { const now = performance.now(), memory = process.memoryUsage(); maxTimerGapMs = Math.max(maxTimerGapMs, now - last); last = now; timerTicks++;
    maxSampledHeapBytes = Math.max(maxSampledHeapBytes, memory.heapUsed); maxSampledExternalBytes = Math.max(maxSampledExternalBytes, memory.external); }, 5);
  const started = performance.now();
  const sink = { async write(bytes) {
    const before = produced; await file.writeFile(bytes);
    // Native compression has a bounded input/output pipeline; one compressed write may overlap source movement.
    if (example && produced !== before) { acknowledgementOverlapWrites++; maxRowsAdvancedDuringAcknowledgement = Math.max(maxRowsAdvancedDuringAcknowledgement, produced - before); }
    firstByteMs ??= performance.now() - started; firstDeliveryProducedRows ??= produced; writes++; outputBytes += bytes.length;
  } };
  let blob, result;
  try {
    const writer = "write" + ({ pdf: "Pdf", xlsx: "Xlsx", csv: "Csv" })[fixture.format];
    if (example) result = await example.writeServiceAnalysis(api, { rows: selectedRows(), sink });
    else if (fixture.delivery === "blob") blob = await api[writer](rows(api, fixture), config);
    else await api[writer + "To"](rows(api, fixture), sink, config);
  } catch (error) { clearInterval(timer); await file.close(); throw error; }
  const engineMs = performance.now() - started, cpu = process.cpuUsage(cpuStart), memoryAfter = process.memoryUsage(), peakRssBytes = process.resourceUsage().maxRSS * 1024;
  await new Promise(resolve => setTimeout(resolve, 6)); clearInterval(timer);
  // Blob copying and all validation are outside the measured interval. Sink file acceptance is inside it.
  if (blob) { const bytes = new Uint8Array(await blob.arrayBuffer()); outputBytes = bytes.length; await file.writeFile(bytes); }
  await file.close();
  if (example && (enumerations !== 1 || returned !== 1 || produced !== fixture.rows || result.rows !== fixture.rows || result.columns !== 6 || result.bytes !== outputBytes))
    throw Error("Editable analysis did not complete its one-pass source and delivery contract.");
  const bytes = await readFile(output);
  return { case: fixture.name, engine, engineMs, firstByteMs, outputBytes, outputSha256: sha(bytes), rows: fixture.rows, columns: titles.length,
    peakRssBytes, maxSampledHeapBytes, maxSampledExternalBytes, memoryBefore, memoryAfter, maxTimerGapMs, timerTicks, cpuMicroseconds: cpu,
    inputScope: example ? "Canonical lazy operationalRows and writeServiceAnalysis included: Unicode sites, typed values/dates, live Excel rules, three formula caches and one visible range chart. No retained source matrix. Imports/setup excluded; current-only diagnostics, no baseline ranking."
      : "Lazy deterministic row generation and portable cell construction included; no retained source matrix. Imports/setup excluded.",
    ...(example ? { sourceContract: { produced, enumerations, returned, writes, firstDeliveryProducedRows, awaitedByteSink: true,
      acknowledgementOverlapWrites, maxRowsAdvancedDuringAcknowledgement, nativeCompressionMayReadAhead: true } } : {}),
    outputScope: fixture.delivery === "sink" ? "Complete export through awaited regular-file writes; OS page cache applies; first byte is first sink acceptance." : "Complete Blob generation; later Blob read and file copy excluded; first-byte metric unavailable.",
    memoryScope: "Fresh Node process lifetime peakRSS through operation end includes import/native memory; sampled JS heap/external are separate and not additive. Sampling can miss synchronous peaks.",
    runtime: source, host: { node: process.version, platform: process.platform, arch: process.arch, release: os.release(), cpus: os.cpus().length, model: os.cpus()[0]?.model,
      loadBefore, loadAfter: os.loadavg(), priority: os.getPriority(), affinity: "not pinned; platform scheduler", pid: process.pid } };
}
async function validate(output = required("output")) {
  const input = await manifest(), fixture = input.cases.find(item => item.name === required("case"));
  if (!fixture) throw Error("Unknown workload.");
  if (sha(await readFile(path.join(path.dirname(fileURLToPath(import.meta.url)), "verify-report-workload.py"))) !== input.verifierSha256) throw Error("Independent verifier changed after preparation.");
  const bytes = await readFile(output); let proof;
  if (fixture.format === "csv") {
    const expected = [titles.join(","), ...Array.from({ length: fixture.rows }, (_, index) => row(index).map(value => value instanceof Date ? value.toISOString() : value).join(","))].join("\r\n") + "\r\n";
    if (!bytes.equals(Buffer.from(expected))) throw Error("CSV differs from every expected value.");
    proof = { rows: fixture.rows, columns: titles.length, everyCell: true };
  } else if (fixture.format === "pdf") {
    const checked = spawnSync(required("qpdf"), ["--check", output], { encoding: "utf8", maxBuffer: 1024 * 1024 });
    if (checked.status !== 0) throw Error("qpdf failed: " + checked.stdout + checked.stderr);
    const extracted = spawnSync(required("pdftotext"), ["-layout", output, "-"], { encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
    if (extracted.status !== 0) throw Error("Poppler extraction failed.");
    const lines = extracted.stdout.split(/[\r\n\f]+/).map(line => line.trim()).filter(line => /^N\d{6}\s/.test(line));
    if (lines.length !== fixture.rows) throw Error("PDF row count differs: " + lines.length);
    for (let index = 0; index < lines.length; index++) if (lines[index].split(/\s+/).join("|") !== row(index).map(value => value instanceof Date ? value.toISOString().slice(0, 10) : value).join("|")) throw Error("PDF cell/order mismatch at row " + index + ": " + lines[index]);
    const text = bytes.toString("latin1"), dataCells = (text.match(/\/S \/TD\b/g) ?? []).length;
    if (fixture.tagged && dataCells !== fixture.rows * titles.length) throw Error("Tagged PDF cell structure count differs: " + dataCells);
    proof = { rows: fixture.rows, columns: titles.length, everyCell: true, qpdf: true, ...(fixture.tagged ? { dataCells } : {}) };
  } else {
    const checked = spawnSync(required("python"), [path.join(path.dirname(fileURLToPath(import.meta.url)), "verify-report-workload.py"), output, String(fixture.rows),
      ...(fixture.kind === "editable-analysis" ? ["service-analysis"] : [])], { encoding: "utf8", maxBuffer: 1024 * 1024 });
    if (checked.status !== 0) throw Error("Independent XLSX validation failed: " + checked.stdout + checked.stderr);
    proof = JSON.parse(checked.stdout);
  }
  return { passed: true, sha256: sha(bytes), bytes: bytes.length, proof };
}

// Functional acknowledgement/abort proof, outside every throughput measurement.
async function contracts() {
  const input = await manifest(), fixture = input.cases.find(item => item.name === required("case"));
  if (fixture?.kind !== "editable-analysis" || !input.example) throw Error("Contract proof requires the canonical editable-analysis case.");
  const source = input.engines.current;
  if (sha(await readFile(source.path)) !== source.sha256) throw Error("Current runtime changed before contract proof.");
  const api = await import(pathToFileURL(source.path).href), example = await import(pathToFileURL(input.example.path).href);
  const output = path.resolve(required("output")), storedOutput = output + ".store.xlsx", file = await open(storedOutput, "w");
  let produced = 0, enumerations = 0, returned = 0, writes = 0, bytes = 0, firstDeliveryProducedRows;
  const progress = [];
  async function* selected() {
    enumerations++;
    try { for (const value of example.operationalRows(fixture.rows)) { produced++; yield value; } }
    finally { returned++; }
  }
  const destination = { async write(chunk) {
    const before = produced; await new Promise(resolve => setTimeout(resolve, 1)); await file.writeFile(chunk);
    if (produced !== before) throw Error("Rows advanced while the slow destination awaited acknowledgement.");
    firstDeliveryProducedRows ??= produced; writes++; bytes += chunk.length;
  } };
  let result;
  try { result = await example.writeServiceAnalysis(api, { input: { compression: "store" }, rows: selected(), sink: destination, onProgress: event => progress.push(event.rows) }); }
  finally { await file.close(); }
  if (enumerations !== 1 || returned !== 1 || produced !== fixture.rows || result.rows !== fixture.rows || result.bytes !== bytes || progress.at(-1) !== fixture.rows)
    throw Error("Slow editable analysis violated source, progress or byte-sink delivery ownership.");
  const storedComplete = await validate(storedOutput);
  // Hold the first entry-data chunk. Native compression may fill its pipeline, but must stop before consuming the full source.
  const compressedFile = await open(output, "w");
  let autoProduced = 0, autoEnumerations = 0, autoReturned = 0, autoWrites = 0, autoBytes = 0, heldFrontier;
  async function* compressedRows() {
    autoEnumerations++;
    try { for (const value of example.operationalRows(fixture.rows)) { autoProduced++; yield value; } }
    finally { autoReturned++; }
  }
  let autoResult;
  try { autoResult = await example.writeServiceAnalysis(api, { rows: compressedRows(), sink: { async write(chunk) {
    if (autoWrites === 1) {
      const before = autoProduced;
      await new Promise(resolve => setTimeout(resolve, 100)); const after100ms = autoProduced;
      await new Promise(resolve => setTimeout(resolve, 100)); const after200ms = autoProduced;
      if (after100ms !== after200ms || after200ms >= fixture.rows) throw Error("Held native-compression output did not stop source consumption at a finite frontier.");
      heldFrontier = { before, after100ms, after200ms, totalRows: fixture.rows, quiescentWhileHeld: true };
    }
    await compressedFile.writeFile(chunk); autoWrites++; autoBytes += chunk.length;
  } } }); } finally { await compressedFile.close(); }
  if (!heldFrontier || autoEnumerations !== 1 || autoReturned !== 1 || autoProduced !== fixture.rows || autoResult.rows !== fixture.rows || autoResult.bytes !== autoBytes)
    throw Error("Compressed editable analysis violated its one-pass or byte-delivery contract.");
  const complete = await validate();
  const controller = new AbortController(), reason = new Error("Cancel the 100k selected analysis after 1000 rows");
  let cancelledRows = 0, cancelledReturns = 0, release, caught; const partial = [];
  const released = new Promise(resolve => { release = resolve; });
  async function* cancelled() {
    try { for (const value of example.operationalRows(fixture.rows)) { if (++cancelledRows === 1000) controller.abort(reason); yield value; } }
    finally { cancelledReturns++; release(); }
  }
  const cancelledDestination = { write(chunk) { partial.push(chunk.slice()); } };
  try { await example.writeServiceAnalysis(api, { input: { compression: "store" }, rows: cancelled(), sink: cancelledDestination, signal: controller.signal }); }
  catch (error) { caught = error; }
  if (caught !== reason || cancelledRows !== 1000) throw Error("Cancellation changed its original reason or row frontier.");
  await new Promise((resolve, reject) => {
    const watchdog = setTimeout(() => reject(Error("Cancelled selected source did not release.")), 10000);
    released.then(() => { clearTimeout(watchdog); resolve(); });
  });
  const partialBytes = Buffer.concat(partial);
  if (cancelledReturns !== 1 || partialBytes.indexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06])) !== -1) throw Error("Cancelled analysis returned a completed archive or did not release its input once.");
  await manifest();
  if (sha(await readFile(source.path)) !== source.sha256) throw Error("Current runtime changed during contract proof.");
  return { passed: true, runtime: source, case: fixture.name, functionalProofNotTiming: true,
    slowSink: { compression: "store", produced, enumerations, returned, writes, bytes, firstDeliveryProducedRows, progress: progress.at(-1), callerOwnedByteSink: true, acknowledgementDelayMs: 1, independentOutput: storedComplete },
    nativeCompression: { produced: autoProduced, enumerations: autoEnumerations, returned: autoReturned, writes: autoWrites, bytes: autoBytes, heldFrontier },
    cancellation: { declaredRows: fixture.rows, consumedRows: cancelledRows, returned: cancelledReturns, originalReason: true, partialBytes: partialBytes.length, completedZip: false },
    independentOutput: complete };
}
try {
  let result;
  if (command === "prepare") {
    const evidence = path.resolve(required("evidence")); await mkdir(evidence, { recursive: true });
    const engines = {};
    for (const engine of ["baseline", "current"]) { const file = path.resolve(required(engine)); engines[engine] = { path: file, sha256: sha(await readFile(file)) }; }
    const names = opts.cases?.split(","), cases = names ? names.map(name => { const item = fixtures.find(item => item.name === name); if (!item) throw Error("Unknown case " + name); return item; }) : fixtures.filter(item => !item.currentOnly);
    const examplePath = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../../OfficeIMO.JavaScript.Examples/report-engine-job.mjs");
    result = { schema: 1, created: new Date().toISOString(), helperSha256: sha(await readFile(fileURLToPath(import.meta.url))),
      verifierSha256: sha(await readFile(path.join(path.dirname(fileURLToPath(import.meta.url)), "verify-report-workload.py"))), engines, cases,
      ...(cases.some(item => item.kind === "editable-analysis") ? { example: { path: examplePath, sha256: sha(await readFile(examplePath)) } } : {}) };
    await writeFile(path.join(evidence, "workloads.json"), JSON.stringify(result, null, 2) + "\n");
  } else if (command === "run") result = await run();
  else if (command === "validate") result = await validate();
  else if (command === "contracts") result = await contracts();
  else if (command === "verify") {
    const input = await manifest();
    for (const engine of Object.values(input.engines)) if (sha(await readFile(engine.path)) !== engine.sha256) throw Error("Measured runtime changed during the cohort.");
    result = { passed: true, engines: input.engines };
  } else throw Error("Use prepare, run, validate, contracts or verify.");
  console.log(JSON.stringify(result));
} catch (error) { console.error(error.stack ?? String(error)); process.exitCode = 1; }
