// Opt-in proof against a caller-selected CanopyX candidate. No production dependency or copied declarations.
import { spawnSync } from "node:child_process";
import { mkdir, writeFile, readFile } from "node:fs/promises";
import { resolve, join } from "node:path";
import { fileURLToPath } from "node:url";
const root = fileURLToPath(new URL("../../OfficeIMO.JavaScript/", import.meta.url));
const output = process.argv[2] && resolve(process.argv[2]), canopy = process.argv[3] && resolve(process.argv[3]);
if (!output || !canopy || !process.env.npm_execpath) throw new Error("Use npm run test:canopy-types -- <packed-evidence-directory> <CanopyX-repository>.");
const archive = JSON.parse(await readFile(join(output, "consumer-report.json"), "utf8")).archive;
const folder = join(output, "canopy-types"); await mkdir(folder, { recursive: true });
await writeFile(join(folder, "package.json"), '{"private":true,"type":"module"}\n');
function run(args) {
  const result = spawnSync(process.execPath, args, { cwd: folder, encoding: "utf8", timeout: 120000 });
  if (result.status !== 0) throw new Error(result.error?.message ?? result.stdout + result.stderr);
}
run([process.env.npm_execpath, "install", "--ignore-scripts", "--no-audit", "--no-fund", "--save-exact", join(output, archive)]);
await writeFile(join(folder, "consumer.mts"), `import type { GridExport, GridHostOptions, GridCursorQuery } from "canopyx/grid";
import { exportCanopy, writeCanopyTo, createCanopyExport } from "@evotecit/officeimo/integrations/canopyx";
const options = { tones: { warning: { background: "FFF2CC" } } };
async function write(capture: GridExport) {
  createCanopyExport(capture, "xlsx", options);
  await exportCanopy(capture, "csv");
  await writeCanopyTo(capture, "xlsx", new WritableStream<Uint8Array>(), options);
}
async function cursor(capture: GridExport<GridCursorQuery>) { await exportCanopy(capture, "pdf", options); }
const host: GridHostOptions = { onExport: async ({ capture, format, signal }) => {
  if (format !== "xlsx" && format !== "csv" && format !== "pdf") throw new Error("Unsupported format");
  await writeCanopyTo(capture, format, new WritableStream<Uint8Array>(), { ...options, signal });
} };
void [write, cursor, host];
`);
await writeFile(join(folder, "tsconfig.json"), JSON.stringify({ compilerOptions: {
  strict: true, exactOptionalPropertyTypes: true, noUncheckedIndexedAccess: true, noEmit: true, target: "ES2022",
  module: "NodeNext", moduleResolution: "NodeNext", lib: ["ES2022", "DOM"],
  paths: { "canopyx/*": [join(canopy, "JavaScript", "*.d.ts")] }
}, files: ["consumer.mts"] }, null, 2) + "\n");
run([join(root, "node_modules/typescript/bin/tsc"), "-p", "tsconfig.json"]);
await writeFile(join(output, "canopy-types.json"), JSON.stringify({ passed: true, compiler: "5.9.3", declarationSource: canopy, archive,
  contracts: ["offset capture", "cursor capture", "host callback capture", "cancellation", "packed optional entry"], runtimeDependencies: 0 }, null, 2) + "\n");
console.log("Actual CanopyX offset/cursor captures and export host callback compile against the installed OfficeIMO archive.");
