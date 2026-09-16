# Optional isolated public-page pilot

This opt-in Linux tool fetches bounded public HTTP(S) bytes on the host, then
runs HTML parsing, JavaScript, document conversion and screen/print rendering
inside one rootless Podman container. Its script worker is a child process in
that same container. The container has no network route or host mounts. This is
an internal pilot, not a public untrusted-content API or a claim that arbitrary
sites are safe or compatible.

From the repository root on Linux with .NET 10, rootless Podman, seccomp and
CPU/memory/PID cgroups:

```bash
mkdir -p /tmp/officeimo-public-pilot/{renderer,worker}
dotnet publish OfficeIMO.Html.Runtime.Worker/OfficeIMO.Html.Runtime.Worker.csproj -c Release -f net10.0 -o /tmp/officeimo-public-pilot/worker
dotnet publish Build/Project/HtmlPublicRenderWorker/OfficeIMO.Html.PublicRenderWorker.csproj -c Release -f net10.0 -o /tmp/officeimo-public-pilot/renderer
podman build -f Build/Project/HtmlPublicRenderWorker/Containerfile.linux-x64 -t localhost/officeimo-html-public-render:local /tmp/officeimo-public-pilot
image_id="$(podman image inspect --format '{{.Id}}' localhost/officeimo-html-public-render:local)"
dotnet run --project Build/Project/HtmlPublicPilot/OfficeIMO.Html.PublicPilot.csproj -c Release -- \
  https://wpt.live/css/css-pseudo/first-letter-001-ref.html \
  /tmp/officeimo-public-pilot/wpt-output \
  "sha256:${image_id}" \
  /tmp/officeimo-public-pilot/renderer/OfficeIMO.Html.PublicRenderWorker.dll \
  /tmp/officeimo-public-pilot/worker/OfficeIMO.Html.Runtime.Worker.dll
```

The output directory must be new. The host does not parse page markup or CSS.
Inside isolation, the renderer discovers HTML scripts, stylesheets, images,
fonts and frame documents, then follows frame-static resources, stylesheet
imports and selected CSS URLs. If execution
requests an unsupplied GET resource, the worker can ask the host to fetch it and
restart the offline capture. Every request passes through the same bounded host
broker; rejected hosts and URLs are recorded as skipped resources. Discovery is
limited to 16 rounds and 24 supplied assets. The acquired bytes, redirects,
connected IP addresses and SHA-256 hashes are recorded before rendering.

Append `--resource=URL` for assets outside this discovery path. An external DNS
name must be explicitly approved with `--host=DNS-name`; explicit resource
URLs approve their own host. The generated OCI probe qualifies active,
multi-candidate responsive-picture selection by viewport, `sizes`, and device density, a two-level relative
JavaScript module graph, a scoped import-map graph with prefix mapping and dynamic import,
and static frame-document loading with relative-resource discovery and inert child scripts.
Child-frame execution realms, frame-body capture/rendering, import attributes,
non-GET requests and browser-wide dynamic loading are
not qualified by this pilot. Cookies and credentials are outside this profile.

The broker allows standard HTTP(S) ports and UTF-8 HTML, validates public IPv4
answers before each direct connection, forbids proxy use and HTTPS downgrade,
and caps a fetch at 20 seconds, five redirects and 4 MiB. Acquisition as a
whole allows 32 attempts and 16 MiB. The OCI lease verifies rootless Podman,
seccomp, CPU/memory/PID cgroups, read-only root, no network or mounts, non-root
user, dropped capabilities and no-new-privileges. The full pipeline has a
two-minute host deadline; the container limits are 512 MiB, one CPU and 32 PIDs.
The container image includes fontconfig and DejaVu fonts for deterministic basic
text rendering. The image is addressed by its full SHA-256 ID, and the returned
renderer and script-worker binary hashes must match the published files supplied
to the runner.

Inspect `acquisition.json`, `outcome.json` or `failure.json` before using any
result as corpus evidence. The successful outcome records the verified policy,
image ID, binary and complete published-file-set hashes, trace entries, output
hashes and confirmed container removal. A failed site does not imply a
browser-equivalent capture. Acquisition and preflight failures also write a
phase-labelled `failure.json` to the new output directory. Remove only
the task-owned image and output after recording evidence; keep shared base-image
caches and unrelated Podman state.
