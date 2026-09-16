# OfficeIMO.Html.Runtime.Rendering

This optional package turns one trusted, scripted application state into
independent OfficeIMO outputs. It composes the provider-neutral runtime host,
the HTML renderer, and the PDF adapter. The runtime's scripts, resources and
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
