# Retained application corpus

This bundle qualifies four application workflows through
`HtmlApplicationDocumentWorkflow`: a live form, a long table, an external
classic/module script graph, and cross-document navigation. Each workflow runs
against exact offline resources, reaches a named final state, captures the live
document, closes the runtime, and creates the standard `ScreenPng`, `PrintPdf`,
and `ScreenToPagePdf` outputs.

The fixtures are independently authored OfficeIMO test inputs covered by the
repository MIT license. Each case retains an `acquisition.json` manifest and the
source bytes addressed by SHA-256 under `inputs/`. The external browser never
uses the network: the comparison runner blocks service workers, intercepts every
request, serves only retained direct GET resources, and fails on an unrecorded
URL, method, or request body.

| Case | Final state and OfficeIMO output | Chromium reference |
| --- | --- | --- |
| Forms | Filled text and textarea values, selected option, checked box, readonly and disabled controls, and generated summary. [Screen](forms/officeimo-screen.png), [print](forms/officeimo-print.pdf), [screen-to-page](forms/officeimo-screen-to-page.pdf) | Chromium 816 x 725; OfficeIMO 816 x 770. Native control metrics and font rasterization remain different. MAE 5.4994, RMSE 28.2212. [Evidence](forms/browser/browser-evidence.json) |
| Tables | Caption, `thead`/`tbody`/`tfoot`, row and column spans, intrinsic columns, 44 body rows, and repeated paged sections. [Screen](tables/officeimo-screen.png), [print](tables/officeimo-print.pdf), [screen-to-page](tables/officeimo-screen-to-page.pdf) | Chromium 816 x 1669; OfficeIMO 816 x 1788. Row height, column sizing, and font metrics remain different; all rows and the final total are retained. MAE 17.3581, RMSE 48.5907. [Evidence](tables/browser/browser-evidence.json) |
| External script graph | Classic script order, static module import, retained JSON fetch, dynamic import after interaction, and final total 51. [Screen](external-graph/officeimo-screen.png), [print](external-graph/officeimo-print.pdf), [screen-to-page](external-graph/officeimo-screen-to-page.pdf) | Both screens are 816 x 720. Button styling and font rasterization remain different. MAE 4.6027, RMSE 28.2947. [Evidence](external-graph/browser/browser-evidence.json) |
| Navigation | Link-driven load of a retained document, `replaceState`, `pushState`, final URL `/report/approved`, and final route text. [Screen](navigation/officeimo-screen.png), [print](navigation/officeimo-print.pdf), [screen-to-page](navigation/officeimo-screen-to-page.pdf) | Both screens are 816 x 720. Font and button metrics account for the remaining difference. MAE 1.4456, RMSE 14.8307. [Evidence](navigation/browser/browser-evidence.json) |

Chromium 151.0.7922.34 ran through Playwright 1.62.0.0 and HtmlTinkerX
3.0.1.0 on Windows x64. The three interactive references retain their exact
browser action manifests. Every `browser-evidence.json` records the action
manifest digest, readiness expressions, source hashes, browser/runtime versions,
image hashes, pixel metrics, and an empty blocked-URL list. The runner bounds the
action file to 256 KiB and 64 actions, caps readiness operations at 60 seconds,
and applies one five-minute budget to browser launch, setup, interaction, capture,
and cleanup.

The corpus exposed a PDF projection defect when nested paint crossed a page
edge while its semantic parent remained in bounds. The shared converter now
clips leaf paint at the page surface, skips empty semantic containers, and keeps
transformed or clipped paint that moves into the visible page. The retained form
proves that page-edge controls serialize, and the table proves that continued
rows remain visible and searchable across page fragments.

The four cases pass on .NET 8 and .NET 10. Every PNG decodes, every PDF reopens,
the expected final-state text is searchable, and both PDF intents contain tagged
content. The retained outputs were produced by the Release net10.0 run. The
[manifest](manifest.json) records every retained file's size and SHA-256 digest.

Run the application contract with:

```powershell
$env:OFFICEIMO_APPLICATION_EVIDENCE_DIR = '<new-output-parent>'
dotnet test OfficeIMO.Html.Runtime.Tests/OfficeIMO.Html.Runtime.Tests.csproj `
  -c Release -f net10.0 `
  --filter FullyQualifiedName~BroaderApplicationCorpusRetainsFinalStateAcrossStandardOutputs
```

These results qualify the named trusted-process workflows and their frozen
rendering outputs. They do not establish general browser, JavaScript, DOM, form,
table, module, or navigation compatibility. Public untrusted pages still use the
separate isolated profile; promoting selected cases through that boundary is a
separate gate.
