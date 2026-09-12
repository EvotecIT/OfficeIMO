# OfficeIMO.Html.Runtime

Execute a trusted local scripted document in a persistent, disposable session and
capture independent OfficeIMO documents. The optional worker uses AngleSharp, AngleSharp.Css,
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

Keep a session open when a workflow needs several operations. Globals, listeners
and the live DOM survive between commands; each capture is an independent snapshot:

```csharp
await using var session = await runtime.OpenTrustedAsync(new HtmlScriptRequest {
    Html = reportHtml,
    Timeout = TimeSpan.FromSeconds(10),
    SessionTimeout = TimeSpan.FromMinutes(2)
}, cancellationToken);

var before = await session.CaptureAsync(cancellationToken: cancellationToken);
await session.ExecuteAsync("document.querySelector('#prepare').click()", cancellationToken);
await session.WaitForAsync("window.reportReady === true", cancellationToken);
var title = await session.EvaluateAsync("document.title", cancellationToken);
var after = await session.CaptureAsync(cancellationToken: cancellationToken);
```

`OpenTrustedAsync` loads the document and runs the supplied scripts. It leaves
readiness to `WaitForAsync` or `CaptureAsync`. Evaluation returns a detached
`JsonElement` using JavaScript JSON serialization; undefined, cyclic values and
other unsupported results fail the session. Captures never expose interpreter objects.

The initial profile targets .NET 8 and .NET 10 hosts and workers. It
supports inline classic scripts, supplied post-load scripts, DOM changes, provider
events, promises and timers. Readiness is an explicit JavaScript expression that
must return boolean `true`. Capture runs in the same event-loop task as that check.
It does not infer network idle, font readiness or layout stability.

Commands are serialized. `Timeout` includes time waiting for another command,
execution and result transfer. A queued cancellation or timeout leaves the active
command and session usable. Cancellation or failure after admission terminates
the worker; open a new session to continue. `SessionTimeout` also counts idle time.
The one-shot `CaptureTrustedAsync` applies `Timeout` to the complete open-and-capture
operation. Disposing a session interrupts active execution and releases the worker.

Unhandled script, timer and listener errors, and promise rejections still unhandled
at a checked script-turn boundary, fail the session instead of returning a successful
snapshot. Rejections handled in the same turn are allowed within
`MaxPendingPromiseRejections`. Listener registration supports callback identity,
object listeners, capture and `once`; passive and signal-controlled registrations
are rejected. Event properties and inline attributes support replacement and removal.

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
isolation, modules, navigation, fetched resources,
and framework-application qualification remain separate runtime work.

Run the standalone report example after building the worker:

```sh
dotnet build OfficeIMO.Html.Runtime.Worker -c Release -f net8.0
dotnet run --project OfficeIMO.Html.Runtime.Examples -c Release -f net8.0 -- OfficeIMO.Html.Runtime.Worker/bin/Release/net8.0/OfficeIMO.Html.Runtime.Worker.dll output/scripted-report
```

The example dispatches a click, waits for promise and timer work to populate a
table, captures the completed report, and writes HTML, Markdown, PNG, SVG and a
searchable PDF. The source fixture and example are independent of HtmlTinkerX.
