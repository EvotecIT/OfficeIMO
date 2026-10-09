// Opt-in comparison assets only. Nothing downloaded here is packed or used by product builds.
import { readFile, writeFile, mkdir } from "node:fs/promises";
import { resolve, join } from "node:path";
import { fileURLToPath } from "node:url";
import { createHash } from "node:crypto";
const directory = fileURLToPath(new URL("./", import.meta.url));
if (!process.argv[2]) throw new Error("Pass a task-owned comparison asset directory.");
const destination = resolve(process.argv[2]), manifestPath = join(directory, "comparison-assets.json");
const manifest = JSON.parse(await readFile(manifestPath, "utf8"));
await mkdir(destination, { recursive: true });
for (const asset of manifest.assets) {
  if (!/^[a-z0-9.-]+\.js$/.test(asset.name)) throw new Error("Invalid comparison asset name.");
  const response = await fetch(asset.url, { signal: AbortSignal.timeout(60000) });
  if (!response.ok) throw new Error(asset.url + ": HTTP " + response.status);
  const bytes = Buffer.from(await response.arrayBuffer()), hash = createHash("sha256").update(bytes).digest("hex");
  if (process.argv.includes("--refresh-hashes")) asset.sha256 = hash;
  else if (asset.sha256 !== hash) throw new Error("Comparison asset hash differs: " + asset.name);
  await writeFile(join(destination, asset.name), bytes);
}
if (process.argv.includes("--refresh-hashes")) await writeFile(manifestPath, JSON.stringify(manifest, null, 2) + "\n");
await writeFile(join(destination, "manifest.json"), JSON.stringify(manifest, null, 2) + "\n");
console.log("Verified " + manifest.assets.length + " pinned test-only comparison assets.");
