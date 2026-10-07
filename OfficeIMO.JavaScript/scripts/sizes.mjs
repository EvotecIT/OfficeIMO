// Explicit evidence measurement: each ESM entry plus its transitive graph, counted once per lane.
import { readFile, writeFile } from "node:fs/promises";
import { gzipSync } from "node:zlib";
import { fileURLToPath } from "node:url";
import { resolve, dirname, join } from "node:path";
const root = fileURLToPath(new URL("../", import.meta.url));
const sizes = [];
for (const layer of ["core", "zip", "xml", "opc", "xlsx", "csv", "pdf", "integrations/datatables"]) {
  const visited = new Map();
  async function visit(path) {
    if (visited.has(path)) return;
    const source = await readFile(path, "utf8"); visited.set(path, source);
    for (const match of source.matchAll(/(?:import|export)\s+[^;]+?from\s+"(\.[^"]+)";/g)) await visit(resolve(dirname(path), match[1]));
  }
  const entry = join(root, "dist", layer, "index.js"); await visit(entry);
  const bytes = Buffer.from([...visited.values()].join("\n"));
  sizes.push({ entry: "/" + layer, modules: visited.size, bytes: bytes.length, gzip9: gzipSync(bytes, { level: 9 }).length,
    perModuleGzip9: [...visited.values()].reduce((n, s) => n + gzipSync(s, { level: 9 }).length, 0) });
}
for (const bundle of ["officeimo", "officeimo-xlsx", "officeimo-csv", "officeimo-pdf", "officeimo-datatables"]) {
  const bytes = await readFile(join(root, "bundles", bundle + ".js"));
  sizes.push({ entry: bundle + ".js", bytes: bytes.length, gzip9: gzipSync(bytes, { level: 9 }).length });
}
const report = { node: process.version, method: "gzip level 9, unminified tsc output; ESM graph concatenated for comparison and separately compressed module sum for HTTP delivery", sizes };
if (process.argv[2]) await writeFile(resolve(process.argv[2]), JSON.stringify(report, null, 2) + "\n");
console.log(JSON.stringify(report, null, 2));
