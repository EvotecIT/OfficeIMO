# Optional isolated public-page pilot

This opt-in CLI exercises the public `HtmlIsolatedPublicPageWorkflow`. It fetches bounded public HTTP(S) bytes on the host, then
runs HTML parsing, JavaScript, document conversion and screen/print rendering
inside one rootless Podman container. Its script worker is a child process in
that same container. The container has no network route or host mounts. This is
a named-page evidence tool, not a claim that arbitrary sites are compatible.

From the repository root on Linux x64 with .NET 10, rootless Podman, seccomp
and CPU/memory/PID cgroups:

```bash
mkdir -p /tmp/officeimo-public-pilot/{renderer,worker}
dotnet publish OfficeIMO.Html.Runtime.Worker/OfficeIMO.Html.Runtime.Worker.csproj -c Release -f net10.0 -r linux-x64 --self-contained false -o /tmp/officeimo-public-pilot/worker
dotnet publish Build/Project/HtmlPublicRenderWorker/OfficeIMO.Html.PublicRenderWorker.csproj -c Release -f net10.0 -r linux-x64 --self-contained false -o /tmp/officeimo-public-pilot/renderer
podman build --platform linux/amd64 -f Build/Project/HtmlPublicRenderWorker/Containerfile.linux -t localhost/officeimo-html-public-render:local /tmp/officeimo-public-pilot
image_id="$(podman image inspect --format '{{.Id}}' localhost/officeimo-html-public-render:local)"
dotnet run --project Build/Project/HtmlPublicPilot/OfficeIMO.Html.PublicPilot.csproj -c Release -- \
  https://wpt.live/css/css-pseudo/first-letter-001-ref.html \
  /tmp/officeimo-public-pilot/wpt-output \
  "sha256:${image_id}" \
  /tmp/officeimo-public-pilot/renderer/OfficeIMO.Html.PublicRenderWorker.dll \
  /tmp/officeimo-public-pilot/worker/OfficeIMO.Html.Runtime.Worker.dll \
  --license=BSD-3-Clause \
  --scenario=wpt-first-letter-reference
```

On Apple Silicon macOS, install Podman and start its rootless AppleHV machine:

```bash
brew install podman
podman machine init --rootful=false
podman machine start
```

Use `linux-arm64` for both container payload publishes and build the same
multi-architecture container file with `--platform linux/arm64`. The host pilot
still runs with `dotnet run`; the default `podman` command connects to the
active machine. If the caller does not inherit the Homebrew path, set
`PodmanCommand` or `--podman-command` to `/opt/homebrew/bin/podman`.

On Windows, keep acquisition in the Windows process and invoke the qualified
rootless Podman installation through WSL by appending:

```text
--podman-command=wsl.exe --podman-arg=-d --podman-arg=Ubuntu --podman-arg=--exec --podman-arg=podman
```

The output directory must be new. The host does not parse page markup or CSS.
Inside isolation, the renderer discovers HTML scripts, stylesheets, images,
fonts and frame documents, then follows frame-static resources, stylesheet
imports and selected CSS URLs. If execution
requests an unsupplied dynamic resource, the worker can ask the host to acquire its
exact URL, method, allowed headers, body, fetch options and occurrence, then restart
the offline capture with that response. A top-level document navigation uses its
own exact request and redirect transcript, including initiator, reduced referrer,
navigation kind, selected history entry, method/body, and occurrence. URL-only static assets retain their simpler
GET path. Every request passes through the same bounded host
broker; rejected hosts and URLs are recorded as skipped resources. Discovery is
limited to 16 rounds and 24 supplied assets. The acquired bytes, redirects,
connected IP addresses and SHA-256 hashes are recorded before rendering.

Append `--resource=URL` for assets outside this discovery path. An external DNS
name must be explicitly approved with `--host=DNS-name`; explicit resource
URLs approve their own host. The generated OCI probe qualifies active,
multi-candidate responsive-picture selection by viewport, `sizes`, and device density, a two-level relative
JavaScript module graph, a scoped import-map graph with prefix mapping and dynamic import,
script-driven relative `fetch()` GET replay with query preservation and fragment stripping,
script-driven asynchronous XMLHttpRequest with request headers, bounded same-origin
GET/HEAD and caller-authorized non-GET request replay through the same transport,
bounded same-origin dynamic redirects, and exact cross-origin requests with
caller-approved origins and CORS preflight,
and same-origin frame-document loading with relative-resource discovery, isolated classic-script child realms,
frame-local import maps and module graphs, static and dynamic JSON imports, bounded parent/child messaging,
and searchable frame-body rendering in both PDF modes. Invalid import attributes and JSON MIME types
reject with `TypeError`; malformed JSON module source rejects with `SyntaxError`.
Cross-origin frame execution, origin-changing dynamic redirects,
transferable frame messages, import-attribute module types beyond JSON and browser-wide
dynamic loading are not qualified by this profile. Cookies and credentials are outside this profile.
GET and HEAD are admitted by default. Add `--method=POST` (or another supported
method) only when the named page is authorized to perform that live request; the
broker executes each discovered occurrence once and records request names, byte
counts and digests without retaining header values or request bodies. Add
`--dynamic-origin=https://api.example/` to authorize one exact cross-origin
dynamic origin. `--host` alone never authorizes a dynamic cross-origin request.
Likewise, `--dynamic-origin` does not approve static assets or document redirects
on that host; use `--host` when those are intended.
Same-origin top-level GET navigation is admitted automatically. Add
`--navigation-origin=https://reports.example/` for each additional exact requested
or redirect origin. Add `--navigation-method=POST` only when the page is authorized
to submit a live navigation body. Navigation authority remains separate from
`--host` and `--dynamic-origin`. Each redirect target is revalidated and must be
authorized; HTTPS downgrade and credentials are rejected. Origin-changing dynamic
fetch redirects remain outside the profile.
Use `--max-output-bytes=BYTES` or `--max-total-output-bytes=BYTES` to lower the
8 MiB per-output or 12 MiB combined encoded-output ceilings.
Use `--timeout-seconds=SECONDS` to lower the acquisition and execution deadline
when exercising cancellation. Verified container removal has a separate fixed
one-minute fail-safe budget and is recorded in failure evidence. `--retain-input` writes acquired bytes only when the
source license and retention policy permit it.

The broker allows standard HTTP(S) ports and UTF-8 HTML, validates public IPv4
answers before each direct connection, forbids proxy use and HTTPS downgrade,
and caps a fetch at 20 seconds, five redirects and 4 MiB. Acquisition as a
whole allows 32 attempts and 16 MiB. The OCI lease verifies rootless Podman,
seccomp, CPU/memory/PID cgroups, read-only root, no network or mounts, non-root
user, dropped capabilities and no-new-privileges. Acquisition and execution have
a two-minute default deadline; verified cleanup may use its separate one-minute
budget. The container limits are 512 MiB, one CPU and 32 PIDs.
The container image includes fontconfig and the repository's manifest-bound
portable browser font pack. Its seven font files and three license files are
validated on the host and inside isolation; the complete package digest and
source manifest are retained in `outcome.json`. The image is addressed by its
full SHA-256 ID, and the returned renderer, script-worker, font-package, and
published-directory identities must match the supplied payloads.

Inspect `acquisition.json`, `outcome.json` or `failure.json` before using any
result as corpus evidence. The successful outcome records the verified policy,
image ID, binary and complete published-file-set hashes, trace entries, output
hashes, font and license provenance, action results, and confirmed container removal. A failed site does not imply a
browser-equivalent capture. Acquisition and preflight failures also write a
phase-labelled `failure.json` to the new output directory. Remove only
the task-owned image and output after recording evidence; keep shared base-image
caches and unrelated Podman state.
