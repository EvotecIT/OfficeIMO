# OfficeIMO.Html.Runtime.Rendering

This optional package provides two application-to-document workflows. The
trusted workflow turns one caller-approved scripted application state into
independent OfficeIMO outputs. The isolated public-page workflow acquires a
named HTTP(S) page through a bounded host broker and runs the complete parser,
JavaScript, capture and rendering pipeline in a verified networkless OCI
container. Both workflows compose the provider-neutral runtime host, the HTML
renderer, and the PDF adapter. The trusted workflow's scripts, resources and
network policy still come from `HtmlScriptRequest`. The workflow's external
resource resolver serves only explicitly supplied responses and responses retained
by the completed capture, with observed responses taking precedence. Data URLs
and system fonts follow the selected render options. The live page and
context close before image or PDF generation begins.
Render-only supplied resources still obey the page's allowed-origin and
redirect limits. A direct URL identity wins over a redirect alias.

Choose explicit output intents. `ScreenFullPage` keeps screen CSS in a continuous
image, `PrintPaged` applies print CSS and pagination, and `ScreenSnapshotPaged`
places screen CSS onto PDF pages. The output also retains the resolved drawing
scene, diagnostics, provider identity, action results, bounded trace, render
resources, and frozen document for other OfficeIMO consumers. Given an
`IHtmlRuntimeHost host` and
prepared `reportHtml` and `reportResources` values:

```csharp
using OfficeIMO.Html;
using OfficeIMO.Html.Pdf;
using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Runtime.Rendering;

HtmlApplicationDocumentResult result = await HtmlApplicationDocumentWorkflow.RunAsync(
    host,
    new HtmlApplicationDocumentRequest {
        Page = new HtmlScriptRequest {
            Profile = HtmlRuntimeProfile.WebApplicationV1,
            DocumentUrl = new Uri("https://reports.example/index.html"),
            Html = reportHtml,
            Resources = reportResources,
            ReadyExpression = "document.querySelector('#total')?.textContent==='Total: 42'"
        },
        Actions = new[] {
            new HtmlAutomationRequest {
                Query = HtmlLocatorQuery.ByAccessibleName("Region"),
                Action = HtmlAutomationAction.SelectOptions,
                Values = new[] { "South" }
            },
            new HtmlAutomationRequest {
                Query = HtmlLocatorQuery.ByText("Total: 18"),
                Action = HtmlAutomationAction.Wait,
                WaitState = HtmlLocatorWaitState.Text,
                Value = "Total: 18"
            }
        },
        FinalReadyExpression = "document.querySelector('#total')?.textContent==='Total: 18'",
        RenderRequests = new[] {
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenFullPage, HtmlRenderEncoder.Png,
                new HtmlRenderOptions { ViewportWidth = 816 }),
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.PrintPaged, HtmlRenderEncoder.Pdf,
                new HtmlToPdfOptions { ViewportWidth = 816 }),
            HtmlRenderRequest.Create(HtmlRenderIntentProfile.ScreenSnapshotPaged, HtmlRenderEncoder.Pdf,
                new HtmlToPdfOptions { ViewportWidth = 816 })
        }
    }, cancellationToken);

byte[] screenPng = result.Outputs[0].Images[0].Bytes;
byte[] printPdf = result.Outputs[1].Pdf!.ToBytes();
byte[] screenToPagePdf = result.Outputs[2].Pdf!.ToBytes();
```

Every output in one application workflow inherits the page request's
`DevicePixelRatio`. Runtime resource capture and static screen, print, and
screen-to-page rendering therefore select responsive images at the same density.
Same-origin captured frame bodies are also projected automatically. The static
renderer lays each child document out in its iframe viewport, applies the selected
screen or print media context, clips overflow, and keeps child text searchable in
PDF output. Root and child capture snapshots remain separate and immutable.

The host can be `HtmlProcessRuntimeProvider` or another provider advertising
the requested profile, structured actions and operation tracing. The current
process worker executes explicitly trusted content only; it is not an OS
sandbox for arbitrary public-site scripts. Supply offline resources or a
deliberate runtime resource policy. The workflow never grants a render pass
fresh network access, and an unavailable resource appears in the
output diagnostics rather than being fetched again. The capture manifest covers
observed resources only; `RenderResources` exposes the supplied and observed
resources retained for rendering. For PDF outputs, the host
resource policy permits the workflow's retained-resource resolver to supply HTTP(S)
responses; it does not add an HTTP client.

## Isolated public pages

`HtmlIsolatedPublicPageWorkflow` is the separate untrusted-content entry point.
It never passes public markup or scripts to `OpenTrustedAsync` on the host. The
host acquires only explicitly admitted public HTTP(S) bytes; the OCI worker has
no network route or host mounts and owns parsing, scripting, capture, screen
rendering, browser-print rendering, and screen-to-page rendering.

```csharp
HtmlIsolatedPublicPageResult result = await HtmlIsolatedPublicPageWorkflow.RunAsync(
    new HtmlIsolatedPublicPageExecutionOptions {
        ImageId = "sha256:<full-image-id>",
        PublishedRendererAssemblyPath = rendererAssembly,
        PublishedWorkerAssemblyPath = workerAssembly,
        // On Windows, invoke the qualified rootless Podman installation in WSL:
        PodmanCommand = "wsl.exe",
        PodmanCommandArguments = new[] { "-d", "Ubuntu", "--exec", "podman" }
    },
    new HtmlIsolatedPublicPageRequest {
        ScenarioId = "wpt-first-letter-reference",
        Url = new Uri("https://wpt.live/css/css-pseudo/first-letter-001-ref.html"),
        SourceLicense = "BSD-3-Clause"
    }, cancellationToken);

ReadOnlyMemory<byte> screenPng = result.Outputs.Single(
    output => output.Name == "screen.png").Content;
```

The current `NetworklessRootlessOciV1` profile requires rootless Podman,
seccomp, CPU/memory/PID cgroups, a read-only root filesystem, no network, no
mounts, UID 65532, dropped capabilities, and no-new-privileges. It verifies the
full immutable image ID, container inspection, renderer and script-worker entry
assembly hashes, complete published-directory hashes, and container removal.
The qualified hosts are Linux with direct Podman, Windows with Podman in WSL2,
and Apple Silicon macOS with a rootless AppleHV Podman machine. On macOS the
default `podman` command uses the active machine connection; callers launched
without the Homebrew path can set `PodmanCommand` to
`/opt/homebrew/bin/podman`.

The acquisition broker permits standard HTTP(S) ports and public IPv4 only,
revalidates DNS at every redirect, disables proxies, cookies, credentials and
decompression, and rejects TLS downgrade. One response is limited to 4 MiB,
the run to 32 acquisition attempts and 16 MiB, and isolated discovery to 16
rounds and 24 supplied resources. Every stricter limit supplied through
`Runtime.ResourcePolicy` is preserved for acquisition and isolated replay.
`OperationTimeout` covers acquisition through output validation; verified
container removal then has a separate fixed one-minute fail-safe budget. Input bytes are omitted from results by
default; set `RetainInputBytes` only when the source license and retention policy
permit it. The result always carries request/final URLs, connected addresses,
redirects, byte counts, SHA-256 digests, runtime trace summaries, known
unsupported features, and output digests. Canceled and failed runs throw typed
exceptions with partial acquisition, phase, worker-identity, trace and cleanup
evidence.

The inner provider ID may be `officeimo.trusted-process`: that process is
trusted by the isolated controller inside the container. It does not describe
the host boundary. `IsolationProfile`, `IsolationPolicy`, immutable image and
payload identities, and confirmed removal describe the untrusted-content
boundary. A successful named-page result proves that page and profile only; it
does not claim Chromium parity or arbitrary-site compatibility.
