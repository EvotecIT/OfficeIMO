# OfficeIMO.Html.Runtime

Execute a trusted local scripted document in a disposable process and capture an
independent OfficeIMO document. The optional worker uses AngleSharp, AngleSharp.Css,
AngleSharp.Js and Jint; ordinary HTML parsing and conversion do not depend on it.

Build or publish `OfficeIMO.Html.Runtime.Worker` and deploy its complete output
directory. Supply its DLL path and the inert DOM services to the process provider:

```csharp
using OfficeIMO.Html.Runtime;
using OfficeIMO.Html.Providers;

var runtime = new HtmlProcessRuntimeProvider(workerDllPath, AngleSharpDomServices.Instance);
var capture = await runtime.CaptureTrustedAsync(new HtmlScriptRequest {
    Html = "<p id='status'>Loading</p><script>setTimeout(() => { document.querySelector('#status').textContent = 'Ready'; window.ready = true; }, 20);</script>",
    ReadyExpression = "window.ready === true",
    Timeout = TimeSpan.FromSeconds(10)
}, cancellationToken);

var status = capture.Document.QuerySelector("#status")!.TextContent;
// Use the captured document with HtmlConversionDocument.FromDocument for conversion.
```

The initial profile targets .NET 8 and .NET 10 hosts and workers. It
supports inline classic scripts, supplied post-load scripts, DOM changes, provider
events, promises and timers. Readiness is an explicit JavaScript expression that
must return boolean `true`. Capture runs in the same event-loop task as that check.
It does not infer network idle, font readiness or layout stability.

Capture transfers nodes and attributes structurally, including namespaces,
document mode and template contents. It does not serialize and reparse HTML.
The result is frozen, has fresh document-local identities and no original source
positions, and remains usable after the worker exits. Editing a captured document
does not resume JavaScript execution. The current capture does not include event
listeners, JavaScript globals, live form-property state, stylesheet changes made
only through CSSOM, or shadow roots. DOM-backed style text and attributes are captured.

The process is terminated on cancellation, timeout or response-budget failure.
This is a **trusted-content execution profile**, not an OS sandbox for hostile
scripts. Network loading and host CLR capabilities are not configured. Strict OS
isolation, persistent interactive sessions, modules, navigation, fetched resources,
and framework-application qualification remain separate runtime work.

Run the standalone report example after building the worker:

```sh
dotnet build OfficeIMO.Html.Runtime.Worker -c Release -f net8.0
dotnet run --project OfficeIMO.Html.Runtime.Examples -c Release -f net8.0 -- OfficeIMO.Html.Runtime.Worker/bin/Release/net8.0/OfficeIMO.Html.Runtime.Worker.dll output/scripted-report
```

The example dispatches a click, waits for promise and timer work to populate a
table, captures the completed report, and writes HTML, Markdown, PNG, SVG and a
searchable PDF. The source fixture and example are independent of HtmlTinkerX.
