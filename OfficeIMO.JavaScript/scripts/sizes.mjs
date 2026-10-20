// Explicit evidence measurement: each ESM entry plus its transitive graph, counted once per lane.
import { readFile, writeFile, readdir } from "node:fs/promises";
import { gzipSync } from "node:zlib";
import { createHash } from "node:crypto";
import { fileURLToPath } from "node:url";
import { resolve, dirname, join } from "node:path";
const root = fileURLToPath(new URL("../", import.meta.url));
const manifest = JSON.parse(await readFile(join(root, "package.json"), "utf8"));
const sizes = [];
for (const [name, target] of Object.entries(manifest.exports)) {
  if (typeof target !== "object" || typeof target.import !== "string") continue;
  const visited = new Map();
  async function visit(path) {
    if (visited.has(path)) return;
    const source = await readFile(path, "utf8"); visited.set(path, source);
    for (const match of source.matchAll(/(?:import|export)\s+[^;]+?from\s+"(\.[^"]+)";/g)) await visit(resolve(dirname(path), match[1]));
  }
  const entry = resolve(root, target.import); await visit(entry);
  const bytes = Buffer.from([...visited.values()].join("\n"));
  sizes.push({ entry: name === "." ? "/" : name.slice(1), modules: visited.size, bytes: bytes.length, sha256: createHash("sha256").update(bytes).digest("hex"), gzip9: gzipSync(bytes, { level: 9 }).length,
    perModuleGzip9: [...visited.values()].reduce((n, s) => n + gzipSync(s, { level: 9 }).length, 0) });
}
for (const bundle of (await readdir(join(root, "bundles"))).filter(name => /^officeimo(?:-[a-z]+)?\.(?:m?js)$/.test(name)).sort()) {
  const bytes = await readFile(join(root, "bundles", bundle));
  sizes.push({ entry: bundle, bytes: bytes.length, sha256: createHash("sha256").update(bytes).digest("hex"), gzip9: gzipSync(bytes, { level: 9 }).length });
}
const report = { node: process.version, method: "gzip level 9, unminified tsc output; ESM graph concatenated for comparison and separately compressed module sum for HTTP delivery", sizes };
if (process.argv[2]) await writeFile(resolve(process.argv[2]), JSON.stringify(report, null, 2) + "\n");
console.log(JSON.stringify(report, null, 2));
