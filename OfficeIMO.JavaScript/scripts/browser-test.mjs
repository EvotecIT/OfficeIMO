// Keep browser installation/session ownership in the existing HtmlTinkerX validation runner.
import { spawn } from "node:child_process";
import { createServer } from "node:http";
import { readFile, mkdir } from "node:fs/promises";
import { resolve, join, relative } from "node:path";
import { fileURLToPath } from "node:url";
const root = fileURLToPath(new URL("../", import.meta.url)), repository = resolve(root, "..");
if (!process.argv[2]) throw new Error("Pass a task-owned evidence directory to npm run test:browser -- <directory>.");
const output = resolve(process.argv[2]); await mkdir(output, { recursive: true });
const configuredDll = process.argv.slice(3).find(argument => argument.startsWith("--interop-dll="));
const interopDll = configuredDll === undefined ? undefined : resolve(configuredDll.slice("--interop-dll=".length));
const argumentsForRunner = process.argv.slice(3).filter(argument => !argument.startsWith("--interop-dll="));
const shapingAssets = argumentsForRunner.find(argument => argument.startsWith("--shaping-assets="));
const shapingRoot = argumentsForRunner.includes("--shaping") && shapingAssets ? resolve(shapingAssets.slice("--shaping-assets=".length)) : undefined;
function run(args) {
  return new Promise((accept, reject) => {
    const child = spawn(process.platform === "win32" ? "dotnet.exe" : "dotnet", args, { cwd: repository, stdio: "inherit" });
    child.once("error", reject); child.once("exit", code => code === 0 ? accept() : reject(new Error("Browser validation exited " + code)));
  });
}
const server = createServer(async (request, response) => {
  try {
    if (new URL(request.url, "http://127.0.0.1").pathname === "/host.html") {
      response.setHeader("Content-Type", "text/html; charset=utf-8"); response.end("<!doctype html><meta charset=utf-8><title>OfficeIMO module qualification</title>"); return;
    }
    const requested = decodeURIComponent(new URL(request.url, "http://127.0.0.1").pathname);
    const providerAsset = shapingRoot !== undefined && requested.startsWith("/__shaping__/");
    const fontAsset = shapingRoot !== undefined && /^\/__fonts__\/(?:Carlito-Regular|NotoSansArabic-Regular)\.ttf$/.test(requested);
    const base = providerAsset ? shapingRoot : fontAsset ? join(repository, "Website/Apps/OfficeIMO.Web.Converter/Assets/Fonts") : root;
    const local = providerAsset ? requested.slice("/__shaping__/".length) : fontAsset ? requested.slice("/__fonts__/".length) : "." + requested;
    const path = resolve(base, local);
    if (relative(base, path).startsWith("..") || !(providerAsset ? /\.(?:js|mjs|wasm)$/ : fontAsset ? /\.ttf$/ : /\.(?:js|mjs)$/).test(path)) { response.writeHead(404).end(); return; }
    response.setHeader("Content-Type", path.endsWith(".wasm") ? "application/wasm" : fontAsset ? "font/ttf" : "text/javascript; charset=utf-8"); response.setHeader("Access-Control-Allow-Origin", "*");
    response.end(await readFile(path));
  } catch { response.writeHead(404).end(); }
});
await new Promise(resolve => server.listen(0, "127.0.0.1", resolve));
try {
  if (!process.argv.includes("--pdf-rich") && !process.argv.includes("--execution") && !process.argv.includes("--shaping"))
    await run(["run", "--project", "OfficeIMO.JavaScript.Examples/OfficeIMO.JavaScript.Examples.csproj", "-c", "Release", "--artifacts-path", join(output, "build"), "--", join(output, "example")]);
  const runner = interopDll === undefined ? ["run", "--project", "Build/BrowserExports/OfficeIMO.Browser.Interop.csproj", "-c", "Release", "--artifacts-path", join(output, "build"), "--"] : [interopDll];
  await run([...runner, repository, join(output, "interop"), "--limits", "--example=" + join(output, "example"),
    "--module-url=http://127.0.0.1:" + server.address().port + "/dist", ...argumentsForRunner]);
} finally { await new Promise(resolve => server.close(resolve)); }
