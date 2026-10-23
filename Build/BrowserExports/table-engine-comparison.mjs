// One complete operation per request. The shared PowerForge runner owns repetition.
import { readFile, writeFile, mkdir, readdir, lstat, realpath } from "node:fs/promises";
import { createHash } from "node:crypto";
import { createRequire } from "node:module";
import { fileURLToPath } from "node:url";
import { createInterface } from "node:readline";
import path from "node:path";
import os from "node:os";
import { packages, fixtures, loadAdapter, exportBuffer, lineSpacing } from "./table-engine-adapters.mjs";

const command = process.argv[2], args = process.argv.slice(3), options = {};
for (let index = 0; index < args.length; index += 2) {
  if (!args[index]?.startsWith("--") || args[index + 1] === undefined) throw Error("Expected named argument/value pairs.");
  const key = args[index].slice(2); if (Object.hasOwn(options, key)) throw Error("Duplicate --" + key);
  options[key] = args[index + 1];
}
const required = name => { if (!options[name]) throw Error("Missing --" + name); return options[name]; };
const sha = bytes => createHash("sha256").update(bytes).digest("hex");
const here = path.dirname(fileURLToPath(import.meta.url));
const contains = (parent, child) => { const relative = path.relative(parent, child); return relative === "" || !relative.startsWith(".." + path.sep) && relative !== ".." && !path.isAbsolute(relative); };
async function fingerprint(file) { const bytes = await readFile(file); return { path: file, bytes: bytes.length, sha256: sha(bytes) }; }

async function moduleGraph(entry) {
  const root = path.dirname(entry), seen = new Map();
  async function visit(file) {
    if (seen.has(file)) return;
    if (!contains(root, file)) throw Error("Runtime import escaped its explicit compiled root: " + file);
    const bytes = await readFile(file); seen.set(file, { path: file, bytes: bytes.length, sha256: sha(bytes) });
    for (const match of bytes.toString("utf8").matchAll(/(?:^|\n)\s*(?:import|export)\s+[^;]*?\bfrom\s+["']([^"']+)["']/g)) {
      if (match[1].startsWith("node:")) continue;
      if (!match[1].startsWith(".")) throw Error("Runtime has an undeclared production import: " + match[1]);
      await visit(path.resolve(path.dirname(file), match[1]));
    }
  }
  await visit(entry);
  const files = [...seen.values()].sort((left, right) => left.path.localeCompare(right.path));
  return { files, sha256: sha(JSON.stringify(files)) };
}

async function snapshotDirectory(root) {
  const files = [], links = [];
  async function visit(directory) {
    for (const name of (await readdir(directory)).sort()) {
      const file = path.join(directory, name), entry = await lstat(file);
      if (entry.isSymbolicLink()) {
        // npm's command links are not package inputs. Other links could escape isolation.
        if (path.basename(directory) === ".bin") { links.push(path.relative(root, file)); continue; }
        throw Error("Dependency snapshot contains an unqualified link: " + file);
      }
      if (entry.isDirectory()) await visit(file);
      else if (entry.isFile()) { const item = await fingerprint(file); files.push({ ...item, path: path.relative(root, file) }); }
      else throw Error("Dependency snapshot contains a special file: " + file);
    }
  }
  await visit(root); return { sha256: sha(JSON.stringify(files)), files, ignoredCommandLinks: links };
}

async function prepare() {
  const evidence = path.resolve(required("evidence")), dependencies = await realpath(required("dependencies"));
  const repository = await realpath(path.resolve(here, "../.."));
  if (contains(repository, dependencies)) throw Error("Comparison dependencies must be in a private directory outside the checkout.");
  const packageJson = path.join(dependencies, "package.json"), packageLock = path.join(dependencies, "package-lock.json");
  const declaration = JSON.parse(await readFile(packageJson, "utf8"));
  if (declaration.private !== true) throw Error("Comparison dependency package must declare private:true.");
  const require = createRequire(packageJson), names = required("engines").split(",");
  if (!names.includes("officeimo") || new Set(names).size !== names.length) throw Error("Select OfficeIMO and unique comparison engines.");
  const engines = { officeimo: { formats: ["csv", "xlsx", "pdf"], version: "explicit source artifact" } };
  for (const name of names) {
    if (name === "officeimo") continue;
    const definition = packages[name]; if (!definition) throw Error("Unknown engine: " + name);
    const members = [definition, ...(definition.plugin ? [definition.plugin] : [])], installed = [];
    for (const member of members) {
      const file = path.join(dependencies, "node_modules", member.package, "package.json"), actual = JSON.parse(await readFile(file, "utf8"));
      if (actual.name !== member.package || actual.version !== member.version) throw Error("Exact comparison package version differs: " + member.package);
      const entry = await realpath(require.resolve(member.package));
      if (!contains(dependencies, entry)) throw Error("Package entry escaped the explicit dependency directory: " + member.package);
      installed.push({ name: actual.name, version: actual.version, license: actual.license, manifest: await fingerprint(file), entry: await fingerprint(entry) });
    }
    engines[name] = { formats: definition.formats, packages: installed, variant: definition.variant ?? name,
      nodeOnly: definition.nodeOnly === true };
  }
  const selected = required("cases").split(",").map(name => {
    const fixture = fixtures.find(item => item.name === name); if (!fixture) throw Error("Unknown case: " + name);
    return { ...fixture, unique: options.unique === "true", textProfile: options["text-profile"] ?? "unicode" };
  });
  if (new Set(selected.map(item => item.name)).size !== selected.length) throw Error("Duplicate comparison cases.");
  if (!["unicode", "bmp"].includes(options["text-profile"] ?? "unicode")) throw Error("Unknown text profile.");
  const licenseReceipt = await fingerprint(path.resolve(required("license-receipt")));
  const approval = JSON.parse(await readFile(licenseReceipt.path, "utf8"));
  if (approval.testOnly !== true || approval.reviewed !== true) throw Error("An explicitly reviewed test-only license receipt is required.");
  for (const engine of Object.values(engines)) for (const item of engine.packages ?? [])
    if (!approval.packages?.some(member => member.name === item.name && member.version === item.version)) throw Error("License receipt does not cover " + item.name + "@" + item.version);
  const officeimoPath = await realpath(required("officeimo"));
  const fonts = {};
  for (const style of ["regular", "bold"]) {
    const file = await realpath(required("font-" + style)), bytes = await readFile(file);
    fonts[style] = { path: file, bytes: bytes.length, sha256: sha(bytes), spacing: lineSpacing(bytes) };
  }
  const input = { schema: 1, createdAtUtc: new Date().toISOString(), evidence, cases: selected, engines,
    officeimo: { ...await fingerprint(officeimoPath), graph: await moduleGraph(officeimoPath) }, dependencies: { root: dependencies, packageJson,
      declaration: await fingerprint(packageJson), lock: await fingerprint(packageLock), snapshot: await snapshotDirectory(path.join(dependencies, "node_modules")), licenseReceipt },
    fonts,
    sources: await Promise.all(["table-engine-comparison.mjs", "table-engine-adapters.mjs"].map(file => fingerprint(path.join(here, file)))),
    contract: "Deterministic DataTables values, complete buffer, fixed widths; frozen font file intake and per-document parsing inside PDF generation. Private comparisons only." };
  await mkdir(evidence, { recursive: true });
  const target = path.join(evidence, "table-engines.json"); await writeFile(target, JSON.stringify(input, null, 2) + "\n");
  return { manifest: target, engines: names, cases: selected.map(item => item.name), dependencyGraphSha256: input.dependencies.snapshot.sha256 };
}

async function manifest() {
  const value = JSON.parse(await readFile(required("manifest"), "utf8"));
  for (const file of [...value.officeimo.graph.files, value.dependencies.declaration, value.dependencies.lock, value.dependencies.licenseReceipt,
    ...Object.values(value.fonts), ...value.sources, ...Object.values(value.engines).flatMap(engine => (engine.packages ?? []).flatMap(item => [item.manifest, item.entry]))])
    if (sha(await readFile(file.path)) !== file.sha256) throw Error("Comparison input changed: " + file.path);
  return value;
}

const sessions = new Map();
async function run(input, request) {
  const spec = input.cases.find(item => item.name === request.case), engine = input.engines[request.engine];
  if (!spec || !engine || !engine.formats.includes(spec.format)) throw Error("Unsupported engine/case selection.");
  if (engine.capabilitySkip && request.probe !== true) throw Error(engine.capabilitySkip);
  if (!["fresh", "cold", "warm"].includes(request.mode)) throw Error("Unknown measurement mode.");
  const fonts = { regularPath: input.fonts.regular.path, boldPath: input.fonts.bold.path,
    regularSpacing: input.fonts.regular.spacing, boldSpacing: input.fonts.bold.spacing };
  const key = request.engine + "/" + request.case;
  let adapter = sessions.get(key), importMs = 0;
  if (request.mode !== "cold" && !adapter) {
    const before = performance.now(); adapter = await loadAdapter(request.engine, input); importMs = performance.now() - before;
    if (request.mode === "warm") sessions.set(key, adapter);
  }
  const memoryBefore = process.memoryUsage(), loadBefore = os.loadavg(), cpuBefore = process.cpuUsage();
  const apiStarted = performance.now();
  if (request.mode === "cold") {
    const before = performance.now(); adapter = await loadAdapter(request.engine, input); importMs = performance.now() - before;
  }
  let last = performance.now(), maxTimerGapMs = 0, sampledHeapBytes = process.memoryUsage().heapUsed;
  const timer = setInterval(() => { const now = performance.now(); maxTimerGapMs = Math.max(maxTimerGapMs, now - last); last = now;
    sampledHeapBytes = Math.max(sampledHeapBytes, process.memoryUsage().heapUsed); }, 5);
  let bytes, engineMs, apiToCompleteMs, cpu, memoryAfter, peakRssBytes;
  try {
    const started = performance.now(); bytes = await exportBuffer(adapter, spec, fonts);
    if (!(bytes instanceof Uint8Array)) throw Error("Adapter did not return a complete binary buffer.");
    engineMs = performance.now() - started; apiToCompleteMs = performance.now() - apiStarted;
    cpu = process.cpuUsage(cpuBefore); memoryAfter = process.memoryUsage(); peakRssBytes = process.resourceUsage().maxRSS * 1024;
    maxTimerGapMs = Math.max(maxTimerGapMs, performance.now() - last);
  } finally { clearInterval(timer); }
  if (bytes.byteLength > 512 * 1024 * 1024) throw Error("Comparison output exceeded the explicit 512 MiB bound.");
  const output = path.resolve(request.output);
  if (!contains(input.evidence, output) || output === input.evidence) throw Error("Output escaped the explicit evidence directory.");
  await mkdir(path.dirname(output), { recursive: true }); await writeFile(output, bytes, { flag: "wx" });
  return { case: spec.name, engine: request.engine, engineVariant: engine.variant ?? request.engine, nodeOnly: engine.nodeOnly === true,
    mode: request.mode, format: spec.format, rows: spec.rows, columns: spec.columns,
    unique: spec.unique, textProfile: spec.textProfile, output, outputBytes: bytes.byteLength, outputSha256: sha(bytes),
    engineMs, importMs, apiToCompleteMs, firstDeliveryMs: request.mode === "cold" ? apiToCompleteMs : engineMs,
    cpuMicroseconds: cpu, memoryBefore, memoryAfter, peakRssBytes, sampledHeapBytes, maxTimerGapMs,
    memoryScope: request.mode === "warm" ? "Persistent process lifetime peak includes earlier exports; not a per-export peak or fresh-process memory ranking."
      : "Fresh process lifetime peak through complete buffer includes imports/input/native allocations; not process-tree memory.",
    timingScope: request.mode === "cold" ? "Public import plus first full buffer export, including frozen PDF font file intake; later artifact write excluded."
      : request.mode === "warm" ? "Exact module/JIT state persists across shared-runner warmups and requests; fresh document/rows/font intake included."
      : "Module preloaded in a fresh process; row/document/font construction and complete buffer export included; JIT is not prewarmed.",
    packageGraphSha256: input.dependencies.snapshot.sha256, runtime: input.officeimo,
    host: { node: process.version, platform: process.platform, arch: process.arch, release: os.release(), cpus: os.cpus().length,
      cpuModel: os.cpus()[0]?.model, priority: os.getPriority(), loadBefore, loadAfter: os.loadavg(), affinity: "not pinned; platform scheduler", pid: process.pid } };
}

async function session() {
  const input = await manifest(); console.log('{"ready":true}');
  for await (const line of createInterface({ input: process.stdin, crlfDelay: Infinity })) {
    try {
      const request = JSON.parse(line);
      if (request.command === "close") break;
      if (request.command !== "run") throw Error("Expected run or close.");
      // One request is one export; there is no hidden warmup or measurement loop.
      console.log(JSON.stringify({ ok: true, result: await run(input, request) }));
    } catch (error) { console.log(JSON.stringify({ ok: false, error: error.stack ?? String(error) })); }
  }
}

try {
  if (command === "prepare") console.log(JSON.stringify(await prepare()));
  else if (command === "session") await session();
  else if (command === "run" || command === "probe") console.log(JSON.stringify(await run(await manifest(), {
    engine: required("engine"), case: required("case"), output: required("output"), mode: required("mode"), probe: command === "probe"
  })));
  else if (command === "verify") {
    const input = await manifest(), current = await snapshotDirectory(path.join(input.dependencies.root, "node_modules"));
    if (current.sha256 !== input.dependencies.snapshot.sha256) throw Error("Installed dependency graph changed during the cohort.");
    console.log(JSON.stringify({ passed: true, dependencyGraphSha256: current.sha256, officeimo: input.officeimo }));
  } else throw Error("Use prepare, session, run, probe or verify.");
} catch (error) { console.error(error.stack ?? String(error)); process.exitCode = 1; }
