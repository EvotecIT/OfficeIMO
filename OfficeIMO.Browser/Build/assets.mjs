// Deterministic concatenation of native JavaScript, like CanopyX's asset assembly.
// No transpilation, minifier, bundler, dependency resolution or publishing.
import { readFile, writeFile, mkdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { join } from "node:path";

const root = fileURLToPath(new URL("../", import.meta.url));
const check = process.argv.includes("--check");
const common = ["common"];
const xlsx = ["xml", "zip", "styles", "xlsx"];
const configurations = [
  ["officeimo-csv", [...common, "csv"], ["writeCsv", "saveBlob"]],
  ["officeimo-xlsx", [...common, ...xlsx], ["createWorkbook", "saveBlob"]],
  ["officeimo", [...common, "csv", ...xlsx], ["createWorkbook", "writeCsv", "saveBlob"]]
];
await mkdir(join(root, "Assets"), { recursive: true });
for (const [name, sources, exports] of configurations) {
  const source = (await Promise.all(sources.map(s => readFile(join(root, "JavaScript/source", s + ".js"), "utf8"))))
    .map(s => s.replace(/\r\n/g, "\n").trimEnd()).join("\n\n");
  const banner = "// Generated: Build/assets.mjs. MIT.\n";
  const files = [
    [name + ".mjs", banner + source + "\n\nexport { " + exports.join(", ") + " };\n"],
    [name + ".js", banner + '(function (root) {\n"use strict";\n' + source +
      "\nObject.assign(root.OfficeIMO || (root.OfficeIMO = {}), { " + exports.join(", ") + " });\n})(globalThis);\n"]
  ];
  for (const [file, content] of files) {
    const path = join(root, "Assets", file);
    if (check) {
      if (await readFile(path, "utf8") !== content) throw new Error("Regenerate " + file + " with npm run build.");
    } else await writeFile(path, content, "utf8");
  }
}
