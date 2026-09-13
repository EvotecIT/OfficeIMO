# OfficeIMO.Html.Runtime

Execute a trusted local scripted document in a persistent, disposable session and
capture independent OfficeIMO documents. The optional worker uses AngleSharp, AngleSharp.Css,
retained AngleSharp.Js DOM bindings and Jint; ordinary HTML parsing and conversion
do not depend on it. The worker builds pinned source dependencies with mutation
and engine-configuration hooks; see [DOM provider provenance](../OfficeIMO.Html.Runtime.AngleSharpDom/README.md)
and [binding provenance](../OfficeIMO.Html.Runtime.AngleSharpJs/README.md).

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
await session.Locator("#prepare").ClickAsync(cancellationToken);
await session.WaitForAsync("window.reportReady === true", cancellationToken);
var title = await session.EvaluateAsync("document.title", cancellationToken);
var after = await session.CaptureAsync(cancellationToken: cancellationToken);
```

`OpenTrustedAsync` loads the document and runs the supplied scripts. It leaves
readiness to `WaitForAsync` or `CaptureAsync`. Evaluation returns a detached
`JsonElement` using JavaScript JSON serialization; undefined, cyclic values and
other unsupported results fail the session. Captures never expose interpreter objects.

Use locators to inspect and interact without composing JavaScript strings:

```csharp
var name = session.Locator(HtmlLocatorQuery.ByAccessibleName("Report name"));
await name.FillAsync("Quarterly", cancellationToken);
await session.Locator(HtmlLocatorQuery.ByText("Add adjustment")).ClickAsync(cancellationToken);
await name.WaitForValueAsync("Quarterly", cancellationToken);
HtmlRuntimeElementState state = await name.InspectAsync(cancellationToken);
```

Locators resolve the current DOM on every operation, including after an application
replaces a node. CSS, normalized text and bounded accessible-name queries support
scopes and explicit `Nth` selection. Actions require exactly one match; `CountAsync`
counts all matches. Text queries select the smallest matching elements. Name queries
use OfficeIMO's shared ARIA, HTML label, alternative-text and title rules, with text
fallback for buttons, links, headings, options and explicitly role-bearing elements.
This is a bounded naming subset, not a complete browser accessibility tree.

The current interaction profile supports DOM clicks, fill for text-like inputs and
textareas, checkbox/radio checking, exact option-value selection, focus, blur and
state waits. Focus events, `document.activeElement`, `:focus` and `:focus-within`
share session state. Fill sends cancelable `beforeinput`, changes the live value,
then sends `input`; a changed text control sends `change` when it loses focus.
`SetCheckedAsync` leaves an already matching checked state untouched, including
indeterminate presentation. Option values must identify unique options; an empty
selection clears them. Disabled, readonly, hidden and inert markup affect readiness.

Missing or temporarily unavailable targets are retried within the session command
deadline. Ambiguous, unsupported and page-rejected actions throw
`HtmlAutomationException` with a structured result and leave the session usable.
Use `AutomateAsync` to receive that result directly or set `WaitForReady = false`.
Script failures and cancellation after command admission still terminate the worker.

Clicks dispatch DOM events, checkbox/radio activation and same-document fragment
navigation. Script-triggered `.click()` uses the same activation owner, including
cancellation and a recursion guard; it can activate hidden or detached elements
without applying the locator's visibility/focus requirements. Downloads, additional
browsing contexts, cross-document links and form submission/reset report
`Unsupported` after click dispatch. Page handlers may already have changed the
document. Pointer hit testing, scrolling and keyboard input remain unqualified.
`IsHiddenByMarkup` does not measure computed visibility, occlusion or layout stability.

Select controls support handler properties and inline handlers, ordinary property
writes, `item`/`namedItem` lookup, and assignment to `value` or `selectedIndex`.
The options collection also supports `item`, `namedItem` and `selectedIndex`.
A value write chooses the first exact, case-sensitive match, even in a multiple
select; an unmatched value or out-of-range index clears selection. Script property
writes do not synthesize input/change events. Automation and capture read the same
native option state. Indexed option replacement, length writes and complete dynamic
default-selection behavior remain outside the qualified select contract.

Supply external scripts and stylesheets without a web server, or explicitly enable
HTTP resource loading:

```csharp
var request = new HtmlScriptRequest {
    DocumentUrl = new Uri("https://reports.example/monthly/index.html"),
    Html = "<p id='status'>Pending</p><script src='../app.js'></script>",
    Resources = new[] {
        HtmlRuntimeResource.FromText(new Uri("https://reports.example/app.js"),
            "document.querySelector('#status').textContent='Ready'", "text/javascript")
    },
    ResourcePolicy = new HtmlRuntimeResourcePolicy {
        AllowNetwork = false,
        MaxResourceBytes = 4 * 1024 * 1024,
        MaxTotalBytes = 16 * 1024 * 1024,
        MaxRequests = 128
    }
};
```

Supplied resources take precedence over HTTP. Network loading is disabled by
default. When enabled, it permits the document origin and explicitly added
`AllowedOrigins`. Each redirect target is checked before a request is sent.
Document resources use GET, without cookies, host credentials, or the host proxy. Resource
deadlines include concurrency admission, redirects and response reading; byte
limits also apply when the server omits a content length. Request counts and total
loaded bytes accumulate across session commands. A missing, blocked, failed or
oversized document resource fails the session.

Classic scripts can use `fetch` with native JavaScript promises:

```javascript
const response = await fetch('/api/report', { signal: controller.signal });
if (!response.ok) throw new Error('Report unavailable: ' + response.status);
const report = await response.json();
document.querySelector('#total').textContent = String(report.total);
```

Create `controller` with `new AbortController()` when cancellation is needed, or
omit the options argument. Fetch resolves HTTP error statuses as responses and
rejects network, policy and deadline failures. Applications can catch these
rejections and continue using the session. Unhandled rejections still fail it.

The fetch profile buffers each complete response before resolving. It supports
GET, HEAD, POST, PUT, PATCH, DELETE and OPTIONS; string, URLSearchParams,
ArrayBuffer and typed-array request bodies; immutable response headers; status,
URL and redirect metadata; and `text()`, `json()`, `arrayBuffer()` and `clone()`.
A non-null response body can be consumed once; clones have independent consumption
state. Text uses UTF-8 decoding. Abort signals cancel queued or active transport
and reject unread response consumption with the signal's reason.

Fetch uses the document's live base URI and the same resource policy and cumulative
budgets as document loads. `MaxRequestBytes` limits each encoded request body;
`MaxTotalRequestBytes` counts bodies sent across commands, including redirect
replays. Supplied resources answer GET and HEAD; other methods require network
permission. Captures retain completed GET loads, including their response
metadata, without replacing assets with POST results.

Additional allowed origins must also pass CORS response checks. Unsafe cross-origin
methods and headers require preflight permission; preflight requests count toward
the resource budgets. Fetch hides cookie headers and filters cross-origin response
headers. Redirects recheck authority, apply method/body rules and remove explicit
Authorization on an origin change. A cross-origin response redirecting to a
different origin is explicitly unsupported. Supplied redirects for fetch must
provide explicit redirect responses and Location headers.

Supported modes are `cors` and `same-origin`, credentials modes are `omit` and
`same-origin`, and redirect modes are `follow` and `error`. There is no cookie jar,
HTTP cache or host credential inheritance. Constructed Request/Response objects,
response streams, Blob/FormData bodies, no-cors/manual redirects, credentialed
cross-origin requests and compressed responses are outside this profile.
Unsupported fetch options are rejected instead of silently changing their meaning.

`capture.Resources` retains immutable loaded responses, including their requested
and final URLs. Use these responses with OfficeIMO's existing
`HtmlRenderResourceResolver` to render after the worker exits without fetching
again. Convert `capture.CreateStandaloneDocument()` to preserve the effective
`capture.BaseUri`, including a relative base element frozen before a route change.
This independent snapshot writes an absolute base href; `capture.Document` retains
the authored attributes and `capture.DocumentUrl` identifies the current route.
The report example demonstrates this wiring. Capture does not wait for arbitrary outstanding
loads; use an explicit condition that represents the application's readiness.

The initial profile targets .NET 8 and .NET 10 hosts and workers. It
supports inline and external classic scripts, supplied post-load scripts, DOM changes, provider
events, promises and timers. Readiness is an explicit JavaScript expression that
must return boolean `true`. Capture runs in the same event-loop task as that check.
It does not infer network idle, font readiness or layout stability.

Global variables, `window`, `globalThis`, `document.defaultView`, window event
receivers and function timer receivers refer to the same session window. Event
paths through the light DOM retain their initial node identities when a
listener changes the tree; shadow-tree paths are unsupported. Function
timers retain their additional callback arguments. String timer handlers are
unsupported. `queueMicrotask` shares Jint's FIFO promise-job queue; callback errors
fail the session.

Session event-loop operations drain pending promise jobs before their work and
complete jobs queued by native DOM actions before admitting the next command.
Typed waits therefore observe promise-driven application updates without an extra
`ExecuteAsync` call. Native timer, resource and lifecycle tasks also finish their
promise jobs before the event loop starts its next task, even while the caller is
not polling the session. A rejection still unhandled at that boundary is retained
as a session failure; adding a handler in a later timer cannot erase it. Command
deadlines and the session lifetime still bound a nonterminating promise job.

`MutationObserver` supports Document, element, text and fragment targets, subtree changes,
attribute filters, old values, `takeRecords()` and `disconnect()`. Callbacks receive
a JavaScript records array and the observer instance. Inapplicable added/removed
node lists are empty NodeLists. Document observation includes direct children and
root replacement. Notifications join the promise-job queue when a mutation first
requires delivery. Pending observers run in one compound notification in order
of their first pending record; promises queued by a callback follow the remaining
observers in that notification. Detached subtrees retain observation until their
notification checkpoint, including nested detachments and overlapping registrations.
Callback ordering follows the DOM Standard's pending set and target/ancestor
registration traversal; engines using observer creation order can differ.
Secondary HTML documents created by `createHTMLDocument` and their clones
share the creating agent's mutation queue but remain inert for script loading.
Wait for an application condition that includes required observer work before capture.
These contracts cover the tested mutation producers; range operations, parser
lifecycle and shadow-root/slot notification integration need further qualification.

`localStorage` and `sessionStorage` provide independent, initially empty in-memory
areas for each runtime session. Values survive commands and application remounts,
but are discarded when the session exits. They do not share data with another
session or a host browser profile, and there are no cross-window storage events.
Both support string keys/values, named property access, enumeration, removal and
clear. `MaxStorageCharacters` limits the combined UTF-16 key/value length in each
area, defaulting to 1 Mi characters. A rejected write throws `QuotaExceededError`
and leaves the previous value intact.

JavaScript module scripts support document-relative and root-relative URLs,
static and dynamic imports, cyclic graphs, live bindings and `import.meta.url`.
Top-level `await` can wait for promise jobs, timers or fetch completion while the
native event loop continues. Classic external scripts resolve dynamic imports
against their own script URL. Captures remain independent after module execution.

Supply module sources as `HtmlRuntimeResource` values with a JavaScript MIME type,
or enable bounded network loading through `ResourcePolicy`. Modules share the
existing response-byte, request, redirect, origin and concurrency budgets.
Cross-origin modules require CORS approval even when their origin is allowlisted.
File and data URLs, credentialed requests and import attributes are unsupported.
Dedicated worker construction throws `NotSupportedError`; a session owns one interpreter.
`MaxModuleCount` bounds retained module sources, including inline roots and failed
loads, and defaults to 1024. Repeated imports reuse the interpreter's module map;
failed source loads remain failures within that session.

One inline import map can precede all module imports. Exact entries, trailing-slash
prefixes and normalized URL scopes are supported; the longest matching scope or
prefix wins, and null entries block resolution. Multiple maps and maps introduced
after imports begin are rejected. Module script loading and evaluation settle
before their document execution operation completes. Full browser lifecycle timing,
module preload, integrity metadata and `import.meta.resolve` remain unqualified.

Same-document application routing supports `history.pushState`, `replaceState`,
`back`, `forward`, and nonzero `go` traversals. Route changes update the retained
document URL and capture identity without replacing the DOM or interpreter.
Relative fetches and newly executed scripts use the active document base. An
active base element keeps its resolved URL across history rewrites; changing or
removing it updates subsequent URL reads and runtime loads.

History state is copied on insertion and restored independently on traversal.
The state graph supports ordinary objects, sparse arrays, cycles, shared references,
maps, sets, dates, regular expressions, errors with causes, boxed primitives,
BigInt, fixed-length ArrayBuffers, typed arrays and DataViews. Functions, symbols,
proxies, promises, DOM nodes, shared/resizable buffers and other unsupported objects
fail with `DataCloneError`. This is history-state storage; it does not expose a
general `structuredClone` or transferable-object API.

`MaxHistoryEntries` defaults to 128 and retains the first and current entries while
evicting older intermediate entries. `MaxHistoryStateBytes` defaults to 1 MiB per
state; `MaxHistoryTotalStateBytes` defaults to 8 MiB across retained states. These
are estimates that count UTF-16 strings, binary data, array slots and graph overhead.
State graphs have a depth limit of 256. A state that cannot fit fails before its
history update. `MaxPendingHistoryTasks` defaults to 1024 for queued traversals and
fragment notifications.

`location` and `document.location` share the session route. Fragment assignment,
`assign`, `replace`, and uncancelled fragment links fire `popstate`, then a queued
`hashchange`; promise jobs run before that hash notification. Back/forward traversal
is asynchronous. History-created events expose their state or URL pair and are
trusted; script redispatch clears that status. `scrollRestoration` stores the
entry's preference, but scrolling is not implemented. Reload, cross-document
loading, additional browsing contexts and the newer Navigation API remain outside
this profile and are not implied by history support.

The test-only Preact 10.29.8 fixture proves UMD loading, mount/unmount, hook effects,
buffered fetch, state updates, controlled input and select events, storage restoration,
mutation delivery and independent captures converted to Markdown and searchable
PDF. These paths do not establish general framework compatibility or a complete
web-application profile. An observer-driven module report additionally proves
JSON loading, combined mutation/promise ordering, typed updates and captures that
convert after session disposal. Full module lifecycle, cross-document navigation,
layout-driven interaction and general framework compatibility remain unqualified.

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
are rejected. Event properties and inline attributes share session-owned registration
for ordinary and collection-backed elements such as select. Replacing a handler
preserves its listener position; clearing and assigning again creates a new position.
Body/window handler aliases share the window target. Specialized error/beforeunload
callback signatures remain outside the qualified handler profile.

Capture transfers nodes and attributes structurally, including namespaces,
document mode and template contents. It does not serialize and reparse HTML.
The result is frozen, has fresh document-local identities and no original source
positions, and remains usable after the worker exits. Editing a captured document
does not resume JavaScript execution. Input values, checkedness, indeterminate
state, textarea values and option selections are retained in `HtmlElement.FormState`.
They survive cloning, import and policy-normalized conversion; semantic inspection
and PDF/SVG/raster rendering use those captured values. Attributes and textarea
text retain their authored defaults, so ordinary HTML serialization does not
restore live values when reopened. Selected files and text selection are not
captured; a capture with selected files fails explicitly.

The current capture does not include event listeners, JavaScript globals, stylesheet changes made
only through CSSOM, or shadow roots. DOM-backed style text and attributes are captured.

The process is terminated on cancellation, timeout or response-budget failure.
This is a **trusted-content execution profile**, not an OS sandbox for hostile
scripts. Host CLR capabilities are not configured. Resource policy covers the
document resource loader; it is not a sandbox for every capability an interpreter
may expose. The tested module and same-document routing paths do not establish
general framework compatibility. XHR, cross-document loading and strict OS
isolation remain separate runtime work.

The retained interpreter has a known async declaration limitation: in a statement
such as `const before = state, result = await operation()`, an earlier initializer
can run again after the await. Keep such declarations in separate statements for
this profile. This behavior was reproduced in Jint 4.16.0 and 4.16.2 and remains an
application-qualification gap.

Run the standalone report example after building the worker:

```sh
dotnet build OfficeIMO.Html.Runtime.Worker -c Release -f net8.0
dotnet run --project OfficeIMO.Html.Runtime.Examples -c Release -f net8.0 -- OfficeIMO.Html.Runtime.Worker/bin/Release/net8.0/OfficeIMO.Html.Runtime.Worker.dll output/scripted-report
```

The example loads supplied external CSS and JavaScript, dispatches a click, fetches
a supplied JSON report, waits for promise and timer work to populate a table, then
captures the completed report.
It renders from retained resources after disposal and writes HTML with its assets,
Markdown, PNG, SVG and a searchable PDF. The source fixture and example are
independent of HtmlTinkerX.
