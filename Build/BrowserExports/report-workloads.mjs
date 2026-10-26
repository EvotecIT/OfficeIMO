// One complete report operation per process. PowerForge owns repetitions and ordering.
import { readFile, writeFile, mkdir, open } from "node:fs/promises";
import { createHash } from "node:crypto";
import { spawnSync } from "node:child_process";
import { pathToFileURL, fileURLToPath } from "node:url";
import os from "node:os";
import path from "node:path";
import { groupedFixtures, groupedRows, writeGroupedWorkload } from "./report-workloads-grouped.mjs";
import { analyticalFixtures, analyticalSetup, analyticalInput, writeAnalyticalWorkload } from "./report-workloads-analytical.mjs";

const args = process.argv.slice(2), command = args.shift();
const opts = Object.fromEntries(args.flatMap((arg, i) => i % 2 ? [] : [[arg.replace(/^--/, ""), args[i + 1]]]));
const required = name => { if (!opts[name]) throw Error("Missing --" + name); return opts[name]; };
const sha = value => createHash("sha256").update(value).digest("hex");
const fixtures = [
  { name: "inventory-csv-100000-sink", format: "csv", rows: 100000, delivery: "sink", styled: false },
  { name: "wide-inventory-csv-100000-sink", format: "csv", rows: 100000, delivery: "sink", kind: "wide", styled: false },
  { name: "quoted-inventory-csv-100000-sink", format: "csv", rows: 100000, delivery: "sink", styled: false, quote: "all" },
  { name: "quoted-wide-inventory-csv-100000-sink", format: "csv", rows: 100000, delivery: "sink", kind: "wide", styled: false, quote: "all" },
  { name: "inventory-xlsx-10000-sink", format: "xlsx", rows: 10000, delivery: "sink", styled: false },
  { name: "inventory-xlsx-100000-sink", format: "xlsx", rows: 100000, delivery: "sink", styled: false },
  { name: "editable-analysis-xlsx-100000-sink", format: "xlsx", rows: 100000, delivery: "sink", kind: "editable-analysis", currentOnly: true },
  { name: "assessment-xlsx-10000-blob", format: "xlsx", rows: 10000, delivery: "blob", styled: true },
  { name: "inventory-pdf-10000-sink", format: "pdf", rows: 10000, delivery: "sink", styled: false },
  { name: "inventory-pdf-100000-sink", format: "pdf", rows: 100000, delivery: "sink", styled: false },
  { name: "wide-inventory-pdf-10000-sink", format: "pdf", rows: 10000, delivery: "sink", kind: "wide", currentOnly: true },
  { name: "tagged-wide-inventory-pdf-10000-sink", format: "pdf", rows: 10000, delivery: "sink", kind: "wide", tagged: true, currentOnly: true },
  { name: "assessment-pdf-10000-blob", format: "pdf", rows: 10000, delivery: "blob", styled: true },
  { name: "tagged-pdf-10000-sink", format: "pdf", rows: 10000, delivery: "sink", styled: true, tagged: true },
  { name: "tagged-carlito-pdf-10000-sink", format: "pdf", rows: 10000, delivery: "sink", styled: true, tagged: true, embeddedFonts: true },
  { name: "tagged-carlito-wide-inventory-pdf-10000-sink", format: "pdf", rows: 10000, delivery: "sink", kind: "wide", tagged: true, embeddedFonts: true, currentOnly: true },
  { name: "tagged-pdf-100000-sink", format: "pdf", rows: 100000, delivery: "sink", styled: true, tagged: true },
  ...groupedFixtures, ...analyticalFixtures
];
const titles = ["ID", "Asset", "Owner", "Status", "Amount", "Seen"];
const wideTitles = [...titles, "Protocol", "Version", "Region", "CPU", "Memory", "Storage", "Requests", "Errors", "TLS", "Backup", "Ticket", "Policy", "Zone", "Health"];
function row(index) {
  return ["N" + String(index).padStart(6, "0"), "node-" + String(index).padStart(6, "0"), "ops-" + index % 17,
    ["Pass", "Warning", "Critical"][index % 3], index + 0.5, new Date("2026-01-" + String(index % 28 + 1).padStart(2, "0") + "T00:00:00Z")];
}
function wideRow(index) { return [...row(index), ...Array.from({ length: 14 }, (_, column) => "M" + (column + 6) + "-" + String(index).padStart(6, "0"))]; }
function* rows(api, fixture) {
  for (let index = 0; index < fixture.rows; index++) {
    const values = fixture.kind === "wide" ? wideRow(index) : row(index);
    if (fixture.styled) values[3] = new api.ExportCell(values[3], { presentation: { background: ["E2F0D9", "FFF2CC", "FCE4D6"][index % 3], bold: index % 3 === 2 } });
    yield values;
  }
}
function settings(fixture) {
  if (fixture.kind === "wide" && fixture.format === "pdf") return { columns: wideTitles.map(header => ({ header })), pageSize: "A3", orientation: "landscape",
    fontSize: 8, padding: 2, pageNumbers: false, compression: true, columnWidths: wideTitles.map(() => 100),
    ...(fixture.tagged ? { tagged: { language: "en-US" } } : {}),
    wideTable: { mode: "bands", repeatColumns: [0, 1], minimumColumnWidth: 100 },
    limits: { maxTableBands: 4, maxTableWindowRows: 64, maxRows: fixture.rows, maxPages: 2000, maxCells: 1000000, maxOutputBytes: 128000000 },
    formatValue: value => value instanceof Date ? value.toISOString().slice(0, 10) : String(value ?? "") };
  const columns = titles.map((header, index) => ({ header, width: [14, 20, 12, 15, 14, 18][index], ...(index === 4 ? { type: "number", format: "0.0" } : index === 5 ? { type: "date", format: "yyyy-mm-dd" } : {}) }));
  if (fixture.kind === "wide") columns.push(...wideTitles.slice(titles.length).map(header => ({ header })));
  if (fixture.format === "pdf") return { columns, pageSize: "A3", orientation: "landscape", fontSize: 8, padding: 2,
    pageNumbers: false, compression: true, formatValue: value => value instanceof Date ? value.toISOString().slice(0, 10) : String(value ?? ""), ...(fixture.tagged ? { tagged: { language: "en-US" } } : {}),
    ...(fixture.styled ? { alternateRowColor: "F3F6F9", headerPresentation: { bold: true, background: "D9E2F3" } } : {}) };
  if (fixture.format === "xlsx") return { columns, dateMode: "utc", sheet: { name: "Inventory", autoSize: { sampleRows: 100 },
    ...(fixture.styled ? { table: { name: "Assessment", style: "TableStyleMedium2" } } : {}) } };
  return { columns, ...(fixture.quote ? { quote: fixture.quote } : {}) };
}
async function manifest() {
  const value = JSON.parse(await readFile(required("manifest"), "utf8"));
  if (sha(await readFile(fileURLToPath(import.meta.url))) !== value.helperSha256) throw Error("Workload helper changed after manifest preparation.");
  if (sha(await readFile(path.join(path.dirname(fileURLToPath(import.meta.url)), "verify-report-workload.py"))) !== value.verifierSha256) throw Error("Independent verifier changed after preparation.");
  if (value.groupedHelperSha256 && sha(await readFile(path.join(path.dirname(fileURLToPath(import.meta.url)), "report-workloads-grouped.mjs"))) !== value.groupedHelperSha256) throw Error("Grouped workload changed after preparation.");
  if (value.example && sha(await readFile(value.example.path)) !== value.example.sha256) throw Error("Canonical editable-analysis example changed after preparation.");
  for (const source of value.additionalSources ?? []) if (sha(await readFile(source.path)) !== source.sha256) throw Error("Analytical workload input changed: " + source.path);
  return value;
}
async function run() {
  const input = await manifest(), fixture = input.cases.find(item => item.name === required("case")), engine = required("engine");
  if (!fixture || !Object.hasOwn(input.engines, engine)) throw Error("Unknown workload or engine.");
  if (fixture.currentOnly && engine !== "current") throw Error("Analytical reports are current-only diagnostics; the snapshot baseline is not equivalent work.");
  const source = input.engines[engine];
  if (sha(await readFile(source.path)) !== source.sha256) throw Error("Measured runtime changed after preparation.");
  const api = await import(pathToFileURL(source.path).href), config = settings(fixture), output = path.resolve(required("output"));
  const example = fixture.kind === "editable-analysis" ? await import(pathToFileURL(input.example.path).href) : undefined;
  const composed = fixture.kind === "analytical", setup = composed || fixture.embeddedFonts ? await analyticalSetup(fixture) : {};
  if (fixture.embeddedFonts) config.fonts = { regular: await api.PdfFont.create(setup.fonts.regular), bold: await api.PdfFont.create(setup.fonts.bold) };
  const analytical = example || fixture.kind === "grouped-analysis" || composed;
  const file = await open(output, "w"), memoryBefore = process.memoryUsage(), loadBefore = os.loadavg(), cpuStart = process.cpuUsage();
  let outputBytes = 0, firstByteMs = null, last = performance.now(), maxTimerGapMs = 0, timerTicks = 0;
  let produced = 0, enumerations = 0, returned = 0, writes = 0, firstDeliveryProducedRows, acknowledgementOverlapWrites = 0, maxRowsAdvancedDuringAcknowledgement = 0;
  function* selectedRows() {
    enumerations++;
    try { for (const value of example ? example.operationalRows(fixture.rows) : fixture.kind === "analytical" ? analyticalInput(fixture) : groupedRows(fixture)) { produced++; yield value; } }
    finally { returned++; }
  }
  let maxSampledHeapBytes = memoryBefore.heapUsed, maxSampledExternalBytes = memoryBefore.external;
  const timer = setInterval(() => { const now = performance.now(), memory = process.memoryUsage(); maxTimerGapMs = Math.max(maxTimerGapMs, now - last); last = now; timerTicks++;
    maxSampledHeapBytes = Math.max(maxSampledHeapBytes, memory.heapUsed); maxSampledExternalBytes = Math.max(maxSampledExternalBytes, memory.external); }, 5);
  const started = performance.now();
  const sink = { async write(bytes) {
    const before = produced; await file.writeFile(bytes);
    // Native compression has a bounded input/output pipeline; one compressed write may overlap source movement.
    if (analytical && produced !== before) { acknowledgementOverlapWrites++; maxRowsAdvancedDuringAcknowledgement = Math.max(maxRowsAdvancedDuringAcknowledgement, produced - before); }
    firstByteMs ??= performance.now() - started; firstDeliveryProducedRows ??= produced; writes++; outputBytes += bytes.length;
  } };
  let blob, result;
  try {
    const writer = "write" + ({ pdf: "Pdf", xlsx: "Xlsx", csv: "Csv" })[fixture.format];
    if (composed) result = await writeAnalyticalWorkload(api, fixture, { rows: selectedRows(), sink, input: setup });
    else if (example) result = await example.writeServiceAnalysis(api, { rows: selectedRows(), sink });
    else if (fixture.kind === "grouped-analysis") result = await writeGroupedWorkload(api, fixture, { rows: selectedRows(), sink });
    else if (fixture.delivery === "blob") blob = await api[writer](rows(api, fixture), config);
    else await api[writer + "To"](rows(api, fixture), sink, config);
  } catch (error) { clearInterval(timer); await file.close(); throw error; }
  const engineMs = performance.now() - started, cpu = process.cpuUsage(cpuStart), memoryAfter = process.memoryUsage(), peakRssBytes = process.resourceUsage().maxRSS * 1024;
  await new Promise(resolve => setTimeout(resolve, 6)); clearInterval(timer);
  // Blob copying and all validation are outside the measured interval. Sink file acceptance is inside it.
  if (blob) { const bytes = new Uint8Array(await blob.arrayBuffer()); outputBytes = bytes.length; await file.writeFile(bytes); }
  await file.close();
  const columns = fixture.kind === "wide" ? wideTitles.length : fixture.kind === "grouped-analysis" || composed ? 7 : titles.length;
  if (analytical && (enumerations !== 1 || returned !== 1 || produced !== fixture.rows || result.rows !== fixture.rows || result.columns !== columns || result.bytes !== outputBytes))
    throw Error("Analytical report did not complete its one-pass source and delivery contract.");
  const bytes = await readFile(output);
  return { case: fixture.name, engine, engineMs, firstByteMs, outputBytes, outputSha256: sha(bytes), rows: fixture.rows, columns,
    peakRssBytes, maxSampledHeapBytes, maxSampledExternalBytes, memoryBefore, memoryAfter, maxTimerGapMs, timerTicks, cpuMicroseconds: cpu,
    ...(fixture.embeddedFonts ? { fontProfile: "embedded-carlito-regular-bold", fontSetupScope: "Font file intake and PdfFont.create run before EngineMs and NodeCpuMs; process-lifetime peak RSS includes font preparation." } : {}),
    inputScope: composed ? "Canonical lazy analyticalRows and writeAnalyticalReport included: three hidden grouping dimensions, typed values/nulls/UTC dates, ancestor/grand aggregates, composed sections and native XLSX editing rules/chart. No retained source matrix. Imports/font intake excluded; current-only diagnostics, no baseline ranking."
      : fixture.kind === "grouped-analysis" ? "Lazy ordered service/site metrics and public grouped report APIs included. Typed numbers/nulls/UTC dates, all-null first group, numeric group/grand totals; no retained source matrix or group body. Imports/setup excluded; current-only diagnostics, no baseline ranking."
      : example ? "Canonical lazy operationalRows and writeServiceAnalysis included: Unicode sites, typed values/dates, live Excel rules, three formula caches and one visible range chart. No retained source matrix. Imports/setup excluded; current-only diagnostics, no baseline ranking."
      : fixture.kind === "wide" && fixture.format === "csv" ? "Lazy 20-column deterministic inventory with typed numeric and UTC-date values. One CSV record contains all selected fields; no retained source matrix. Imports/setup excluded."
      : fixture.kind === "wide" ? "Lazy 20-column deterministic inventory with two repeated identifiers and 100-point widths. Complete horizontal-band output; " +
        (fixture.equivalentBaselineContract ? "operator-declared matching baseline, independently validated in both lanes. " : "current-only diagnostics, no baseline ranking. ") + "No retained source matrix."
      : "Lazy deterministic row generation and portable cell construction included; no retained source matrix. Imports/setup excluded.",
    ...(analytical ? { sourceContract: { produced, enumerations, returned, writes, firstDeliveryProducedRows, awaitedByteSink: true,
      acknowledgementOverlapWrites, maxRowsAdvancedDuringAcknowledgement, nativeCompressionMayReadAhead: true } } : {}),
    outputScope: fixture.delivery === "sink" ? "Complete export through awaited regular-file writes; OS page cache applies; first byte is first sink acceptance." : "Complete Blob generation; later Blob read and file copy excluded; first-byte metric unavailable.",
    executionScope: "Fresh Node process for this sample; module import precedes EngineMs, but JIT state is fresh. PowerForge warmup exports ran in different processes; this is not steady-state warmed export latency.",
    memoryScope: "Fresh Node process lifetime peakRSS through operation end includes import/native memory; sampled JS heap/external are separate and not additive. Sampling can miss synchronous peaks.",
    runtime: source, host: { node: process.version, platform: process.platform, arch: process.arch, release: os.release(), cpus: os.cpus().length, model: os.cpus()[0]?.model,
      loadBefore, loadAfter: os.loadavg(), priority: os.getPriority(), affinity: "not pinned; platform scheduler", pid: process.pid } };
}
async function validate(output = required("output")) {
  const input = await manifest(), fixture = input.cases.find(item => item.name === required("case"));
  if (!fixture) throw Error("Unknown workload.");
  if (sha(await readFile(path.join(path.dirname(fileURLToPath(import.meta.url)), "verify-report-workload.py"))) !== input.verifierSha256) throw Error("Independent verifier changed after preparation.");
  const bytes = await readFile(output); let proof;
  if (fixture.kind === "analytical") {
    if (fixture.format === "pdf") {
      const checked = spawnSync(required("qpdf"), ["--check", output], { encoding: "utf8", maxBuffer: 1024 * 1024 });
      if (checked.status !== 0) throw Error("qpdf failed: " + checked.stdout + checked.stderr);
    }
    const checked = spawnSync(required("python"), [path.join(path.dirname(fileURLToPath(import.meta.url)), "verify-analytical-workload.py"), output, String(fixture.rows), fixture.format,
      ...(fixture.format === "pdf" ? [required("pdftotext")] : [])], { encoding: "utf8", maxBuffer: 1024 * 1024 });
    if (checked.status !== 0) throw Error("Independent analytical readback failed: " + checked.stdout + checked.stderr);
    return { passed: true, sha256: sha(bytes), bytes: bytes.length, proof: { ...JSON.parse(checked.stdout), ...(fixture.format === "pdf" ? { qpdf: true } : {}) } };
  }
  if (fixture.format === "csv") {
    const csvTitles = fixture.kind === "wide" ? wideTitles : titles, sourceRow = fixture.kind === "wide" ? wideRow : row;
    const expectedField = value => {
      const text = String(value instanceof Date ? value.toISOString() : value);
      return fixture.quote === "all" ? '"' + text.replace(/"/g, '""') + '"' : text;
    };
    const expected = [csvTitles.map(expectedField).join(","), ...Array.from({ length: fixture.rows }, (_, index) => sourceRow(index).map(expectedField).join(","))].join("\r\n") + "\r\n";
    if (!bytes.equals(Buffer.from(expected))) throw Error("CSV differs from every expected value.");
    proof = { rows: fixture.rows, columns: csvTitles.length, everyCell: true };
  } else if (fixture.format === "pdf") {
    const checked = spawnSync(required("qpdf"), ["--check", output], { encoding: "utf8", maxBuffer: 1024 * 1024 });
    if (checked.status !== 0) throw Error("qpdf failed: " + checked.stdout + checked.stderr);
    let structure;
    if (fixture.tagged) {
      const checked = spawnSync(required("python"), [path.join(path.dirname(fileURLToPath(import.meta.url)), "verify-pdf-workload-structure.py"),
        output, String(fixture.rows), String(fixture.kind === "wide" ? wideTitles.length : titles.length),
        ...(fixture.embeddedFonts ? ["--embedded-fonts"] : [])], { encoding: "utf8", maxBuffer: 1024 * 1024 });
      if (checked.status !== 0) throw Error("Independent tagged PDF structure readback failed: " + checked.stdout + checked.stderr);
      structure = JSON.parse(checked.stdout);
    }
    if (fixture.kind === "grouped-analysis") {
      const checked = spawnSync(required("python"), [path.join(path.dirname(fileURLToPath(import.meta.url)), "verify-report-workload.py"), output, String(fixture.rows), "grouped-pdf", String(fixture.groupRows), required("pdftotext")], { encoding: "utf8", maxBuffer: 1024 * 1024 });
      if (checked.status !== 0) throw Error("Independent grouped PDF validation failed: " + checked.stdout + checked.stderr);
      return { passed: true, sha256: sha(bytes), bytes: bytes.length, proof: { ...JSON.parse(checked.stdout), qpdf: true } };
    }
    const extracted = spawnSync(required("pdftotext"), ["-layout", output, "-"], { encoding: "utf8", maxBuffer: 64 * 1024 * 1024 });
    if (extracted.status !== 0) throw Error("Poppler extraction failed.");
    if (fixture.kind === "wide") {
      // A3 landscape has 1,118 printable points: two 100-point IDs plus nine data columns per band.
      // The oracle is independent Poppler text, checked against every source value and complete page windows.
      const bands = [[0, 1, ...Array.from({ length: 9 }, (_, index) => index + 2)], [0, 1, ...Array.from({ length: 9 }, (_, index) => index + 11)]];
      const orders = bands.map(() => []), windows = new Map();
      for (const page of extracted.stdout.split("\f")) {
        const lines = page.split(/[\r\n]+/).map(line => line.trim()).filter(line => /^N\d{6}\s/.test(line));
        if (!lines.length) continue;
        let pageBand = -1;
        const ids = [];
        for (const line of lines) {
          const actual = line.split(/\s+/), index = Number(actual[0].slice(1)), expected = wideRow(index).map(value => value instanceof Date ? value.toISOString().slice(0, 10) : String(value));
          if (index >= fixture.rows) throw Error("Wide PDF invented an input row.");
          const band = bands.findIndex(columns => columns.map(column => expected[column]).join("|") === actual.join("|"));
          if (band < 0 || pageBand >= 0 && pageBand !== band) throw Error("Wide PDF value/order/band differs: " + line);
          pageBand = band; orders[band].push(index); ids.push(index);
        }
        const key = ids.join(","), covered = windows.get(key) ?? new Set();
        if (covered.has(pageBand)) throw Error("Wide PDF duplicated a band/window.");
        covered.add(pageBand); windows.set(key, covered);
      }
      if (orders.some(order => order.length !== fixture.rows || order.some((value, index) => value !== index)) || [...windows.values()].some(covered => covered.size !== bands.length))
        throw Error("Wide PDF omitted a row/identifier or misaligned its horizontal windows.");
      return { passed: true, sha256: sha(bytes), bytes: bytes.length, proof: { rows: fixture.rows, columns: wideTitles.length, bands: bands.length,
        windows: windows.size, everyCell: true, repeatedIdentifiers: true, alignedWindows: true, qpdf: true, poppler: true,
        ...(structure ? { ...structure, singleLogicalCellPerSourceValue: true } : {}) } };
    }
    const lines = extracted.stdout.split(/[\r\n\f]+/).map(line => line.trim()).filter(line => /^N\d{6}\s/.test(line));
    if (lines.length !== fixture.rows) throw Error("PDF row count differs: " + lines.length);
    for (let index = 0; index < lines.length; index++) if (lines[index].split(/\s+/).join("|") !== row(index).map(value => value instanceof Date ? value.toISOString().slice(0, 10) : value).join("|")) throw Error("PDF cell/order mismatch at row " + index + ": " + lines[index]);
    proof = { rows: fixture.rows, columns: titles.length, everyCell: true, qpdf: true, ...(structure ?? {}) };
  } else {
    const checked = spawnSync(required("python"), [path.join(path.dirname(fileURLToPath(import.meta.url)), "verify-report-workload.py"), output, String(fixture.rows),
      ...(fixture.kind === "editable-analysis" ? ["service-analysis"] : fixture.kind === "grouped-analysis" ? ["grouped-analysis", String(fixture.groupRows)] : [])], { encoding: "utf8", maxBuffer: 1024 * 1024 });
    if (checked.status !== 0) throw Error("Independent XLSX validation failed: " + checked.stdout + checked.stderr);
    proof = JSON.parse(checked.stdout);
  }
  return { passed: true, sha256: sha(bytes), bytes: bytes.length, proof };
}

// Functional acknowledgement/abort proof, outside every throughput measurement.
async function contracts() {
  const input = await manifest(), fixture = input.cases.find(item => item.name === required("case"));
  if (!["editable-analysis", "grouped-analysis", "analytical"].includes(fixture?.kind)) throw Error("Contract proof requires an analytical report case.");
  const source = input.engines.current;
  if (sha(await readFile(source.path)) !== source.sha256) throw Error("Current runtime changed before contract proof.");
  const api = await import(pathToFileURL(source.path).href), example = fixture.kind === "editable-analysis" ? await import(pathToFileURL(input.example.path).href) : undefined;
  const setup = fixture.kind === "analytical" ? await analyticalSetup(fixture) : {};
  const operation = options => example ? example.writeServiceAnalysis(api, options) : fixture.kind === "analytical"
    ? writeAnalyticalWorkload(api, fixture, { ...options, input: { ...setup, ...options.input } }) : writeGroupedWorkload(api, fixture, options);
  const selectedValues = () => example ? example.operationalRows(fixture.rows) : fixture.kind === "analytical" ? analyticalInput(fixture) : groupedRows(fixture);
  const output = path.resolve(required("output")), storedOutput = output + ".store." + fixture.format, file = await open(storedOutput, "w");
  let produced = 0, enumerations = 0, returned = 0, writes = 0, bytes = 0, firstDeliveryProducedRows;
  const progress = [];
  async function* selected() {
    enumerations++;
    try { for (const value of selectedValues()) { produced++; yield value; } }
    finally { returned++; }
  }
  const destination = { async write(chunk) {
    const before = produced; await new Promise(resolve => setTimeout(resolve, 1)); await file.writeFile(chunk);
    if (produced !== before) throw Error("Rows advanced while the slow destination awaited acknowledgement.");
    firstDeliveryProducedRows ??= produced; writes++; bytes += chunk.length;
  } };
  let result;
  try { result = await operation({ input: { compression: "store" }, rows: selected(), sink: destination, onProgress: event => progress.push(event.rows) }); }
  finally { await file.close(); }
  if (enumerations !== 1 || returned !== 1 || produced !== fixture.rows || result.rows !== fixture.rows || result.bytes !== bytes || progress.at(-1) !== fixture.rows)
    throw Error("Slow editable analysis violated source, progress or byte-sink delivery ownership.");
  const storedComplete = await validate(storedOutput);
  // Hold output after source data has entered the writer. Compression may fill its pipeline, but must stop before full source consumption.
  const compressedFile = await open(output, "w");
  let autoProduced = 0, autoEnumerations = 0, autoReturned = 0, autoWrites = 0, autoBytes = 0, heldFrontier;
  async function* compressedRows() {
    autoEnumerations++;
    try { for (const value of selectedValues()) { autoProduced++; yield value; } }
    finally { autoReturned++; }
  }
  let autoResult;
  try { autoResult = await operation({ rows: compressedRows(), sink: { async write(chunk) {
    if (!heldFrontier && autoProduced > 0) {
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
  const controller = new AbortController(), reason = new Error("Cancel the selected analysis after 1000 rows");
  let cancelledRows = 0, cancelledReturns = 0, release, caught; const partial = [];
  const released = new Promise(resolve => { release = resolve; });
  async function* cancelled() {
    try { for (const value of selectedValues()) { if (++cancelledRows === 1000) controller.abort(reason); yield value; } }
    finally { cancelledReturns++; release(); }
  }
  const cancelledDestination = { write(chunk) { partial.push(chunk.slice()); } };
  try { await operation({ input: { compression: "store" }, rows: cancelled(), sink: cancelledDestination, signal: controller.signal }); }
  catch (error) { caught = error; }
  if (caught !== reason || cancelledRows !== 1000) throw Error("Cancellation changed its original reason or row frontier.");
  await new Promise((resolve, reject) => {
    const watchdog = setTimeout(() => reject(Error("Cancelled selected source did not release.")), 10000);
    released.then(() => { clearTimeout(watchdog); resolve(); });
  });
  const partialBytes = Buffer.concat(partial);
  const completedOutput = fixture.format === "xlsx" ? partialBytes.indexOf(Buffer.from([0x50, 0x4b, 0x05, 0x06])) !== -1 : partialBytes.includes(Buffer.from("%%EOF"));
  if (cancelledReturns !== 1 || completedOutput) throw Error("Cancelled analysis returned a completed output or did not release its input once.");
  await manifest();
  if (sha(await readFile(source.path)) !== source.sha256) throw Error("Current runtime changed during contract proof.");
  return { passed: true, runtime: source, case: fixture.name, functionalProofNotTiming: true,
    slowSink: { compression: "store", produced, enumerations, returned, writes, bytes, firstDeliveryProducedRows, progress: progress.at(-1), callerOwnedByteSink: true, acknowledgementDelayMs: 1, independentOutput: storedComplete },
    nativeCompression: { produced: autoProduced, enumerations: autoEnumerations, returned: autoReturned, writes: autoWrites, bytes: autoBytes, heldFrontier },
    cancellation: { declaredRows: fixture.rows, consumedRows: cancelledRows, returned: cancelledReturns, originalReason: true, partialBytes: partialBytes.length, completedOutput: false },
    independentOutput: complete };
}
try {
  let result;
  if (command === "prepare") {
    const evidence = path.resolve(required("evidence")); await mkdir(evidence, { recursive: true });
    if (opts.baseline && opts.baselines !== undefined) throw Error("Choose --baseline or --baselines, not both.");
    const baselines = opts.baselines === undefined ? { baseline: required("baseline") } : JSON.parse(await readFile(path.resolve(opts.baselines), "utf8"));
    if (!baselines || typeof baselines !== "object" || Array.isArray(baselines) || !Object.keys(baselines).length)
      throw Error("Named baselines must be a nonempty object mapping labels to bundle paths.");
    const engines = Object.create(null), baselineNames = Object.keys(baselines).sort();
    for (const engine of baselineNames) {
      if (!/^[a-z][a-z0-9-]*$/.test(engine) || engine === "current") throw Error("Baseline labels must be lowercase names other than current.");
      if (typeof baselines[engine] !== "string" || !baselines[engine].trim()) throw Error("A named baseline must provide a bundle path.");
      const file = path.resolve(baselines[engine]); engines[engine] = { path: file, sha256: sha(await readFile(file)) };
    }
    const current = path.resolve(required("current")); engines.current = { path: current, sha256: sha(await readFile(current)) };
    const names = opts.cases?.split(","), selected = names ? names.map(name => { const item = fixtures.find(item => item.name === name); if (!item) throw Error("Unknown case " + name); return item; }) : fixtures.filter(item => !item.currentOnly);
    const equivalentBaselineCases = opts["equivalent-baseline-cases"]?.split(",") ?? [];
    if (new Set(equivalentBaselineCases).size !== equivalentBaselineCases.length) throw Error("Equivalent baseline cases must be unique.");
    for (const name of equivalentBaselineCases) if (!selected.some(item => item.name === name && item.kind === "wide" && item.currentOnly))
      throw Error("Equivalent baseline declarations apply only to explicitly selected current-only wide inventory cases: " + name);
    const cases = selected.map(item => equivalentBaselineCases.includes(item.name) ? { ...item, currentOnly: false,
      equivalentBaselineContract: "Operator declares that the selected baseline implements the same complete horizontal bands, repeated identifiers, logical cells and sink workload; both artifacts require independent readback." } : item);
    const examplePath = path.resolve(path.dirname(fileURLToPath(import.meta.url)), "../../OfficeIMO.JavaScript.Examples/report-engine-job.mjs");
    const additionalPaths = cases.some(item => item.kind === "analytical") ? ["report-workloads-analytical.mjs", "verify-analytical-workload.py",
      "../../OfficeIMO.JavaScript.Examples/analytical-report-job.mjs", "../../OfficeIMO.JavaScript.Examples/report-engine-job.mjs",
      "../../Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-Regular.ttf", "../../Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-Bold.ttf"] : [];
    if (cases.some(item => item.embeddedFonts)) additionalPaths.push("report-workloads-analytical.mjs",
      "../../Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-Regular.ttf", "../../Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts/Carlito-Bold.ttf");
    if (cases.some(item => item.tagged)) additionalPaths.push("verify-pdf-workload-structure.py", "verify-semantic-report.py");
    const additionalSources = await Promise.all([...new Set(additionalPaths)].map(async name => {
      const file = path.resolve(path.dirname(fileURLToPath(import.meta.url)), name); return { path: file, sha256: sha(await readFile(file)) };
    }));
    result = { schema: 1, created: new Date().toISOString(), helperSha256: sha(await readFile(fileURLToPath(import.meta.url))),
      groupedHelperSha256: sha(await readFile(path.join(path.dirname(fileURLToPath(import.meta.url)), "report-workloads-grouped.mjs"))),
      additionalSources,
      verifierSha256: sha(await readFile(path.join(path.dirname(fileURLToPath(import.meta.url)), "verify-report-workload.py"))), engines, cases, equivalentBaselineCases,
      comparisonBaselines: baselineNames, executionMode: "fresh-process-module-preloaded-jit-fresh",
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
