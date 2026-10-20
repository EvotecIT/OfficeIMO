// One resource operation per Node process. PSPublishModule owns repetitions, ordering and reports.
import { readFile, writeFile, mkdir } from "node:fs/promises";
import { createHash } from "node:crypto";
import { deflateSync } from "node:zlib";
import { spawnSync } from "node:child_process";
import os from "node:os";
import path from "node:path";
import { fileURLToPath, pathToFileURL } from "node:url";

const arguments_ = process.argv.slice(2), mode = arguments_.shift();
const options = Object.fromEntries(arguments_.map((value, index) => index % 2 === 0 ? [value.replace(/^--/, ""), arguments_[index + 1]] : null).filter(Boolean));
const hash = bytes => createHash("sha256").update(bytes).digest("hex");
const required = name => { if (!options[name]) throw Error("Missing --" + name); return options[name]; };
const fontDirectory = "Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts";
const fontCases = [
  { name: "font-carlito", kind: "font", relative: fontDirectory + "/Carlito-Regular.ttf", text: "Zażółć gęślą jaźń 42" },
  { name: "font-japanese", kind: "font", relative: fontDirectory + "/NotoSansJP-OfficeIMO-Common.ttf", text: "日本語 42" }
];
const pdfOptions = { pageNumbers: false, compression: false };

async function graphHash(root) {
  const files = new Map();
  async function visit(file) {
    if (files.has(file)) return;
    const bytes = await readFile(path.join(root, file)), source = bytes.toString("utf8"); files.set(file, hash(bytes));
    for (const match of source.matchAll(/(?:^|\n)\s*(?:import|export)\s+[^;]*?\bfrom\s+["']([^"']+)["']/g)) {
      if (!match[1].startsWith(".")) throw Error("Unexpected external dependency in the measured PDF graph.");
      const dependency = path.posix.normalize(path.posix.join(path.posix.dirname(file), match[1]));
      if (dependency.startsWith("../")) throw Error("PDF dependency escaped the compiled root.");
      await visit(dependency);
    }
  }
  await visit("pdf/index.js");
  return hash(JSON.stringify([...files].sort(([left], [right]) => left.localeCompare(right))));
}
async function library(root) { return import(pathToFileURL(path.join(root, "pdf/index.js")).href); }

function pngInput(dimension, Crc32) {
  const rgb = createHash("sha256"), alpha = createHash("sha256");
  const filtered = Buffer.alloc((dimension * 4 + 1) * dimension), rgbRow = Buffer.alloc(dimension * 3), alphaRow = Buffer.alloc(dimension);
  for (let y = 0; y < dimension; y++) {
    const row = y * (dimension * 4 + 1); filtered[row] = 0;
    for (let x = 0; x < dimension; x++) {
      const r = x & 255, g = y & 255, b = (x + y) & 255, a = 64 + ((x + y) % 192), pixel = row + 1 + x * 4;
      filtered[pixel] = r; filtered[pixel + 1] = g; filtered[pixel + 2] = b; filtered[pixel + 3] = a;
      rgbRow[x * 3] = r; rgbRow[x * 3 + 1] = g; rgbRow[x * 3 + 2] = b; alphaRow[x] = a;
    }
    rgb.update(rgbRow); alpha.update(alphaRow);
  }
  const chunk = (type, data) => {
    const name = Buffer.from(type), output = Buffer.alloc(data.length + 12), crc = new Crc32();
    output.writeUInt32BE(data.length); name.copy(output, 4); data.copy(output, 8); crc.update(name); crc.update(data); output.writeUInt32BE(crc.value, data.length + 8); return output;
  };
  const header = Buffer.alloc(13); header.writeUInt32BE(dimension); header.writeUInt32BE(dimension, 4); header[8] = 8; header[9] = 6;
  return { bytes: Buffer.concat([Buffer.from([137,80,78,71,13,10,26,10]), chunk("IHDR", header), chunk("IDAT", deflateSync(filtered)), chunk("IEND", Buffer.alloc(0))]),
    rgbSha256: rgb.digest("hex"), alphaSha256: alpha.digest("hex") };
}

async function exportPdf(api, fixture, bytes, preparedFont) {
  if (fixture.kind === "png") return api.writePdfDocument([{ kind: "image", image: { format: "png", data: bytes } }], pdfOptions);
  return api.writePdfDocument([{ kind: "paragraph", text: fixture.text }], { ...pdfOptions, fonts: { regular: preparedFont ?? new api.PdfFont(bytes) } });
}

// Narrow independent inspection of xref-resolved image objects; no OfficeIMO decoder is used here.
function validateImagePlanes(bytes, fixture) {
  const tail = bytes.subarray(Math.max(0, bytes.length - 2048)).toString("latin1"), start = Number(/startxref\n(\d+)\n%%EOF/.exec(tail)?.[1]);
  if (!Number.isSafeInteger(start)) throw Error("PDF has no valid xref offset.");
  const lines = bytes.subarray(start).toString("latin1").split("\n"), count = Number(lines[1]?.split(" ")[1]);
  if (lines[0] !== "xref" || !Number.isSafeInteger(count)) throw Error("PDF xref table is invalid.");
  const planes = [];
  for (let id = 1; id < count; id++) {
    const offset = Number(lines[id + 2]?.slice(0, 10)), header = bytes.subarray(offset, offset + 2048).toString("latin1");
    const marker = header.indexOf("\nstream\n");
    if (marker < 0 || !/\/Subtype \/Image/.test(header.slice(0, marker))) continue;
    const dictionary = header.slice(0, marker), length = Number(/\/Length (\d+)/.exec(dictionary)?.[1]);
    const width = Number(/\/Width (\d+)/.exec(dictionary)?.[1]), height = Number(/\/Height (\d+)/.exec(dictionary)?.[1]);
    if (width !== fixture.dimension || height !== fixture.dimension || /\/Filter/.test(dictionary)) throw Error("Unexpected image dimensions/filter.");
    const pixels = bytes.subarray(offset + marker + 8, offset + marker + 8 + length), color = /\/ColorSpace \/(\w+)/.exec(dictionary)?.[1];
    const expectedLength = width * height * (color === "DeviceRGB" ? 3 : color === "DeviceGray" ? 1 : 0);
    if (!expectedLength || pixels.length !== expectedLength) throw Error("Image plane byte length differs.");
    const expectedHash = color === "DeviceRGB" ? fixture.rgbSha256 : fixture.alphaSha256;
    if (hash(pixels) !== expectedHash) throw Error("Independent image plane hash differs from source samples.");
    planes.push(color);
  }
  if (planes.length !== 2 || !planes.includes("DeviceRGB") || !planes.includes("DeviceGray")) throw Error("RGB and alpha planes were not both exported.");
  return { planes, pixels: fixture.dimension ** 2 };
}

async function validate(output, fixture, expected) {
  const bytes = await readFile(output), sha256 = hash(bytes);
  if (sha256 !== expected) throw Error("Complete PDF SHA-256 differs from the immutable baseline: " + sha256);
  const native = spawnSync(required("qpdf"), ["--check", output], { encoding: "utf8", maxBuffer: 1024 * 1024 });
  if (native.status !== 0) throw Error("qpdf validation failed: " + native.stdout + native.stderr);
  let proof;
  if (fixture.kind === "png") proof = validateImagePlanes(bytes, fixture);
  else {
    const text = spawnSync(required("pdftotext"), [output, "-"], { encoding: "utf8", maxBuffer: 1024 * 1024 });
    if (text.status !== 0 || text.stdout.replace(/\s/g, "") !== fixture.text.replace(/\s/g, "")) throw Error("Independent font PDF extraction differs.");
    proof = { text: fixture.text };
  }
  return { passed: true, sha256, bytes: bytes.length, qpdf: native.stdout.trim(), proof };
}

async function prepare() {
  const evidence = path.resolve(required("evidence")), repository = path.resolve(required("repository"));
  const roots = { baseline: path.resolve(required("baseline")), current: path.resolve(required("current")) };
  await mkdir(path.join(evidence, "inputs"), { recursive: true });
  const { Crc32 } = await import(pathToFileURL(path.join(roots.current, "zip/index.js")).href);
  let cases = [];
  const dimensions = required("dimensions").split(",").map(Number);
  for (const dimension of dimensions) {
    if (!Number.isSafeInteger(dimension) || dimension < 1 || dimension > 4096) throw Error("PNG dimensions must be integers through 4096.");
  }
  let reused;
  if (options.reuse) {
    const source = await readFile(options.reuse), previous = JSON.parse(source);
    if (await graphHash(roots.baseline) !== previous.graphHashes.baseline) throw Error("Cached fixture baseline graph differs.");
    const names = [...dimensions.map(dimension => "png-" + dimension), ...fontCases.map(fixture => fixture.name)];
    cases = names.map(name => { const fixture = previous.cases.find(item => item.name === name); if (!fixture) throw Error("Missing cached fixture: " + name); return fixture; });
    for (const fixture of cases) {
      const bytes = await readFile(fixture.input);
      if (hash(bytes) !== fixture.inputSha256 || fixture.baselineProof?.sha256 !== fixture.expectedSha256 || !fixture.baselineProof?.passed)
        throw Error("Cached fixture/proof changed: " + fixture.name);
    }
    reused = { manifest: path.resolve(options.reuse), sha256: hash(source) };
  } else {
    for (const dimension of dimensions) {
      const png = pngInput(dimension, Crc32), input = path.join(evidence, "inputs", "rgba-" + dimension + ".png");
      await writeFile(input, png.bytes);
      cases.push({ name: "png-" + dimension, kind: "png", dimension, input, inputSha256: hash(png.bytes), inputBytes: png.bytes.length,
        rgbSha256: png.rgbSha256, alphaSha256: png.alphaSha256 });
    }
    for (const fixture of fontCases) {
      const input = path.join(repository, fixture.relative), bytes = await readFile(input);
      cases.push({ ...fixture, input, inputSha256: hash(bytes), inputBytes: bytes.length });
    }
    const baseline = await library(roots.baseline);
    for (const fixture of cases) {
      const blob = await exportPdf(baseline, fixture, new Uint8Array(await readFile(fixture.input))), output = path.join(evidence, "inputs", fixture.name + "-expected.pdf");
      const bytes = new Uint8Array(await blob.arrayBuffer()); await writeFile(output, bytes); fixture.expectedSha256 = hash(bytes); fixture.expectedBytes = bytes.length;
      fixture.baselineProof = await validate(output, fixture, fixture.expectedSha256);
    }
  }
  const manifest = { schema: 1, created: new Date().toISOString(), node: process.version, roots,
    graphHashes: { baseline: await graphHash(roots.baseline), current: await graphHash(roots.current) },
    helperSha256: hash(await readFile(fileURLToPath(import.meta.url))), ...(reused ? { reused } : {}), cases };
  await writeFile(path.join(evidence, "inputs.json"), JSON.stringify(manifest, null, 2));
  return { prepared: true, cases: cases.map(fixture => ({ name: fixture.name, bytes: fixture.inputBytes, expectedBytes: fixture.expectedBytes })), manifest: path.join(evidence, "inputs.json") };
}

async function run() {
  const manifest = JSON.parse(await readFile(required("manifest"), "utf8")), fixture = manifest.cases.find(item => item.name === required("case")), engine = required("engine");
  if (!fixture || !manifest.roots[engine]) throw Error("Unknown resource case or engine.");
  const root = manifest.roots[engine], engineGraphSha256 = await graphHash(root);
  if (engineGraphSha256 !== manifest.graphHashes[engine]) throw Error("Engine graph changed after input preparation.");
  const api = await library(root), bytes = new Uint8Array(await readFile(fixture.input));
  if (hash(bytes) !== fixture.inputSha256) throw Error("Fixture input changed.");
  const helperSha256 = hash(await readFile(fileURLToPath(import.meta.url)));
  if (helperSha256 !== manifest.helperSha256) throw Error("Benchmark helper changed after input preparation.");
  const memoryBefore = process.memoryUsage(), loadBefore = os.loadavg(), cpuStart = process.cpuUsage();
  let last = performance.now(), maxTimerGapMs = 0, ticks = 0;
  const timer = setInterval(() => { const now = performance.now(); maxTimerGapMs = Math.max(maxTimerGapMs, now - last); last = now; ticks++; }, 2);
  await new Promise(resolve => setTimeout(resolve, 0));
  const started = performance.now();
  let blob, font;
  try {
    if (fixture.kind === "png") blob = await exportPdf(api, fixture, bytes);
    else font = engine === "current" ? await api.PdfFont.create(bytes) : new api.PdfFont(bytes);
  } catch (error) { clearInterval(timer); throw error; }
  const engineMs = performance.now() - started, cpu = process.cpuUsage(cpuStart), memoryAfter = process.memoryUsage(), peakRssBytes = process.resourceUsage().maxRSS * 1024;
  await new Promise(resolve => setTimeout(resolve, 3)); clearInterval(timer);
  // Font PDF generation, array-buffer copying, hashes, file I/O and independent validation are untimed.
  if (!blob) blob = await exportPdf(api, fixture, bytes, font);
  const output = path.resolve(required("output")), outputBytes = new Uint8Array(await blob.arrayBuffer()); await writeFile(output, outputBytes);
  return { case: fixture.name, engine, engineMs, maxTimerGapMs, timerTicks: ticks, peakRssBytes,
    heapUsedDelta: memoryAfter.heapUsed - memoryBefore.heapUsed, externalDelta: memoryAfter.external - memoryBefore.external,
    arrayBuffersDelta: memoryAfter.arrayBuffers - memoryBefore.arrayBuffers, memoryBefore, memoryAfter,
    cpuMicroseconds: cpu, outputBytes: outputBytes.length, outputSha256: hash(outputBytes), engineGraphSha256, helperSha256, inputSha256: fixture.inputSha256,
    host: { platform: process.platform, release: os.release(), arch: process.arch, cpus: os.cpus().length, model: os.cpus()[0]?.model, loadBefore, loadAfter: os.loadavg(), node: process.version, versions: process.versions,
      priority: os.getPriority(), affinity: "not pinned; platform scheduler", processId: process.pid, execArgv: process.execArgv },
    boundaries: { engineMs: fixture.kind === "font" ? "Font intake only; PDF generation/validation excluded" : "Full PNG decode and uncompressed PDF Blob generation; input read/import excluded",
      peakRssBytes: "Fresh Node process lifetime through measured-operation end; startup/input/native allocations included, later font-PDF generation excluded; no process-tree total",
      maxTimerGapMs: "Largest observed 2 ms interval, including a trailing tick to observe synchronous blocking; no throughput gate",
      memoryDeltas: "Node memoryUsage snapshots; heapUsed is JS heap, external/arrayBuffers are separate overlapping native-backed categories; no sampled peak claim" } };
}

try {
  let result;
  if (mode === "prepare") result = await prepare();
  else if (mode === "run") result = await run();
  else if (mode === "validate") {
    const manifest = JSON.parse(await readFile(required("manifest"), "utf8")), fixture = manifest.cases.find(item => item.name === required("case"));
    if (!fixture) throw Error("Unknown resource case.");
    result = await validate(required("output"), fixture, fixture.expectedSha256);
  } else if (mode === "verify") {
    const manifest = JSON.parse(await readFile(required("manifest"), "utf8")), actual = {};
    for (const engine of ["baseline", "current"]) {
      actual[engine] = await graphHash(manifest.roots[engine]);
      if (actual[engine] !== manifest.graphHashes[engine]) throw Error(engine + " PDF graph changed during the cohort.");
    }
    result = { passed: true, checked: new Date().toISOString(), graphHashes: actual };
  } else throw Error("Use prepare, run, validate or verify.");
  console.log(JSON.stringify(result));
} catch (error) { console.error(error.stack ?? String(error)); process.exitCode = 1; }
