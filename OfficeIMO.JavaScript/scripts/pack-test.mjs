import { spawnSync } from "node:child_process";
import { mkdir, readFile, writeFile, copyFile } from "node:fs/promises";
import { resolve, join } from "node:path";
import { fileURLToPath } from "node:url";

const root = fileURLToPath(new URL("../", import.meta.url));
if (!process.argv[2]) throw new Error("Pass a task-owned evidence directory to npm run test:pack -- <directory>.");
const output = resolve(process.argv[2]), consumer = join(output, "consumer");
await mkdir(consumer, { recursive: true });
const npmCli = process.env.npm_execpath;
if (!npmCli) throw new Error("Run through npm run test:pack so the configured npm CLI is used.");
function run(args, cwd) {
  const result = spawnSync(process.execPath, args, { cwd, encoding: "utf8", timeout: 120000 });
  if (result.status !== 0) throw new Error(result.error?.message ?? result.stdout + result.stderr);
  return result.stdout;
}
const packed = run([npmCli, "pack", "--pack-destination", output, "--json"], root);
const manifest = JSON.parse(packed.slice(packed.indexOf("[")));
for (const file of manifest[0].files) {
  if (!/^(?:package\.json|README\.md|LICENSE|dist\/.+\.(?:js|d\.ts)|bundles\/officeimo(?:-xlsx|-csv|-pdf|-datatables|-canopyx)?\.(?:js|mjs))$/.test(file.path))
    throw new Error("Unexpected shipped file: " + file.path);
}
await writeFile(join(output, "archive.json"), JSON.stringify(manifest, null, 2) + "\n");
await writeFile(join(consumer, "package.json"), '{"private":true,"type":"module"}\n');
for (const file of ["consumer.mts", "worker.mts", "runtime.mjs", "worker-runtime.mjs"]) await copyFile(join(root, "test/consumer", file), join(consumer, file));
run([npmCli, "install", "--ignore-scripts", "--no-audit", "--no-fund", "--save-exact", join(output, manifest[0].filename)], consumer);
run([join(root, "node_modules/typescript/bin/tsc"), "--strict", "--exactOptionalPropertyTypes", "--noUncheckedIndexedAccess", "--noEmit", "--target", "ES2022",
  "--module", "NodeNext", "--moduleResolution", "NodeNext", "--lib", "ES2022,DOM", "consumer.mts"], consumer);
run([join(root, "node_modules/typescript/bin/tsc"), "--strict", "--exactOptionalPropertyTypes", "--noUncheckedIndexedAccess", "--outDir", "worker-runtime", "--target", "ES2022",
  "--module", "NodeNext", "--moduleResolution", "NodeNext", "--lib", "ES2022,WebWorker", "worker.mts"], consumer);
console.log(run(["runtime.mjs"], consumer).trim());
console.log(run(["worker-runtime.mjs"], consumer).trim());
const installed = JSON.parse(await readFile(join(consumer, "package-lock.json"), "utf8"));
if (Object.keys(installed.packages).length !== 2) throw new Error("Packed consumer acquired an unexpected runtime dependency.");
await copyFile(join(consumer, "packed-consumer.xlsx"), join(output, "packed-consumer.xlsx"));
await copyFile(join(consumer, "packed-streamed.xlsx"), join(output, "packed-streamed.xlsx"));
await copyFile(join(consumer, "packed-worker.xlsx"), join(output, "packed-worker.xlsx"));
await copyFile(join(consumer, "packed-consumer.pdf"), join(output, "packed-consumer.pdf"));
await writeFile(join(output, "consumer-report.json"), JSON.stringify({ passed: true, subpaths: 7, strictTypeScript: "5.9.3", archive: manifest[0].filename,
  integrations: ["datatables", "canopyx"], shippedFiles: manifest[0].files.length, runtimeDependencies: 0 }, null, 2) + "\n");
console.log("Strict types, archive contents, isolated npm install and all runtime subpaths passed.");
