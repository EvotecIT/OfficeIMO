# Standard application output API

This evidence was produced through the public `HtmlApplicationDocumentWorkflow`
and `HtmlApplicationOutputOptions` contract. The same request shape executed four
application classes, applied structured interactions, captured the final live DOM,
closed the runtime, and created the named `ScreenPng`, `PrintPdf`, and
`ScreenToPagePdf` outputs.

| Application class | Final state | Retained outputs |
| --- | --- | --- |
| Vanilla ES modules | Review route, Quarterly / South, Total: 25 | [screen](vanilla/officeimo-screen.png), [print](vanilla/officeimo-print.pdf), [screen-to-page](vanilla/officeimo-screen-to-page.pdf) |
| Production React build | Review route, Quarterly / South, Total: 21 | [screen](react-build/officeimo-screen.png), [print](react-build/officeimo-print.pdf), [screen-to-page](react-build/officeimo-screen-to-page.pdf) |
| Preact | Quarterly / South, Total: 18 | [screen](preact/officeimo-screen.png), [print](preact/officeimo-print.pdf), [screen-to-page](preact/officeimo-screen-to-page.pdf) |
| Legacy inline event handler | Approved | [screen](legacy/officeimo-screen.png), [print](legacy/officeimo-print.pdf), [screen-to-page](legacy/officeimo-screen-to-page.pdf) |

The four parameterized cases passed on Windows x64 under .NET 10. Each screen
PNG decoded successfully. Every PDF reopened through `PdfReadDocument`, retained
the expected final application text, and exposed tagged content. The assertions
also verify that all outputs inherit the live page viewport and select the bounded
browser user-agent style profile. The screen artifacts were visually inspected
for the expected final state. [The manifest](manifest.json) records every retained
application output's size and SHA-256 digest.

A separate net8.0 console consumer restored only
`OfficeIMO.Html.Runtime.Rendering` from the task-local feed through an empty,
task-owned NuGet cache. It compiled the standard request and named-result API,
created the supported process host against the deployed worker, and reported
`officeimo.trusted-process:All`. The [package consumer report](package-consumer.json)
records the exact package digest and resolved dependency graph. The empty cache is
material to this proof because a normal machine cache can contain an older package
with the same development version.

The isolated public renderer was rebuilt after adopting the same standard output
API. Its immutable image
`sha256:a337ea0a6cc10b9de61902eaf397874c19ca4fd9400801a8724a7f6c1cf99130`
passed all 22 parser, script, module, fetch/XHR, frame, output-limit, and
cancellation cases plus all six controlled acquisition cases. Every started
container was removed. The retained [OCI summary](oci-summary.json) records the
image and complete published-file digests, isolation policy, discovery rounds,
case outcomes, and cleanup evidence.

Run the evidence-producing contract with:

```powershell
$env:OFFICEIMO_APPLICATION_EVIDENCE_DIR = '<new-output-directory>'
dotnet test OfficeIMO.Html.Runtime.Tests/OfficeIMO.Html.Runtime.Tests.csproj `
  -c Release -f net10.0 --no-restore `
  --filter FullyQualifiedName~OneWorkflowCapturesAndRendersFourApplicationClasses
```

These cases qualify the public composition API and selected runtime behaviors.
They do not establish general React, Preact, JavaScript, DOM, or browser
compatibility. The inputs are trusted and offline; public scripts use the
separate isolated-page workflow.
