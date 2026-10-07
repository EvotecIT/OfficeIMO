// Keep browser installation/session ownership in the existing HtmlTinkerX validation runner.
import { spawn } from "node:child_process";
import { createServer } from "node:http";
import { readFile, mkdir } from "node:fs/promises";
import { resolve, join, relative } from "node:path";
import { fileURLToPath } from "node:url";
const root = fileURLToPath(new URL("../", import.meta.url)), repository = resolve(root, "..");
if (!process.argv[2]) throw new Error("Pass a task-owned evidence directory to npm run test:browser -- <directory>.");
const output = resolve(process.argv[2]); await mkdir(output, { recursive: true });
function run(args) {
  return new Promise((accept, reject) => {
    const child = spawn(process.platform === "win32" ? "dotnet.exe" : "dotnet", args, { cwd: repository, stdio: "inherit" });
    child.once("error", reject); child.once("exit", code => code === 0 ? accept() : reject(new Error("Browser validation exited " + code)));
  });
}
const server = createServer(async (request, response) => {
  try {
    const path = resolve(root, "." + decodeURIComponent(new URL(request.url, "http://127.0.0.1").pathname));
    if (relative(root, path).startsWith("..") || !/\.(?:js|mjs)$/.test(path)) { response.writeHead(404).end(); return; }
    response.setHeader("Content-Type", "text/javascript; charset=utf-8"); response.setHeader("Access-Control-Allow-Origin", "*");
    response.end(await readFile(path));
  } catch { response.writeHead(404).end(); }
});
await new Promise(resolve => server.listen(0, "127.0.0.1", resolve));
try {
  await run(["run", "--project", "OfficeIMO.Browser.Examples/OfficeIMO.Browser.Examples.csproj", "-c", "Release", "--", join(output, "example")]);
  await run(["run", "--project", "Build/BrowserExports/OfficeIMO.Browser.Interop.csproj", "-c", "Release", "--", repository, join(output, "interop"),
    "--limits", "--example=" + join(output, "example"), "--module-url=http://127.0.0.1:" + server.address().port + "/dist", ...process.argv.slice(3)]);
} finally { await new Promise(resolve => server.close(resolve)); }
