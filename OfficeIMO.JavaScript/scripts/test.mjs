// Node 18 on Windows does not expand --test globs; pass explicit test paths.
import { readdir } from "node:fs/promises";
import { fileURLToPath } from "node:url";
import { spawnSync } from "node:child_process";

const directory = new URL("../test/", import.meta.url);
const files = (await readdir(directory)).filter(name => name.endsWith(".test.mjs")).sort();
if (!files.length) throw new Error("No OfficeIMO JavaScript tests found.");
const result = spawnSync(process.execPath, ["--test", ...files.map(name => fileURLToPath(new URL(name, directory)))], { stdio: "inherit" });
if (result.error) throw result.error;
process.exitCode = result.status ?? 1;
