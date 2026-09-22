# Cross-realm window-call checkpoint

OfficeIMO code commits `a7fd998c2137b5deb735fbad237d492aa21d9891` and
`63049d43781ae0b4532a1cce8bc3dae52429f285` qualify a bounded call path
for functions defined on a same-origin child or popup `window`. A parent call
executes in the function's realm with the
parent as the script entry document. Primitive and DOM/window receivers,
arguments and results cross this boundary; unsupported JavaScript objects and
functions raise `NotSupportedError`. Cross-origin or opaque access and calls
into retired realms are rejected.

The child-defined `callDocumentMethod` test adapts the
[WPT synchronous-call case](https://github.com/web-platform-tests/wpt/blob/02c25cc3a34a70a6225147f76cec312483df4706/html/webappapis/dynamic-markup-insertion/opening-the-input-stream/url-entry-document-sync-call.window.js)
for `open`, `write` and `writeln`. This is focused runtime evidence, not an
official WPT harness run or a claim of general cross-realm object compatibility.

The retained provider commits are AngleSharp
`85941d3d35aba79da852f9c64b60e193b9a42206`, AngleSharp.Js
`1190abed8991e0aafb29664462a7177ef64d66d2`, and Jint
`2339be034e48b1f2195987fff27cda225582b7b7`. The bridge itself belongs
to OfficeIMO's session/realm owner and required no provider patch.

From `/tmp`, with the OfficeIMO project path absolute so the installed
10.0.303 SDK is selected, the final focused `dotnet test` filters passed:

| Framework | Filtered suites | Result |
| --- | --- | --- |
| net10.0 | `RuntimeDocumentOpenTests`, `RuntimeAuxiliaryWindowTests`, `RuntimeFrameBaseUrlTests`, `RuntimeResourceTests` | 116 passed, 0 failed |
| net8.0 | `RuntimeDocumentOpenTests`, `RuntimeAuxiliaryWindowTests` | 68 passed, 0 failed |

An independent read-only review found that `.call(...)` lost explicit `this`.
The reviewed fix transfers supported receivers and rejects unsupported ones;
the targeted confirmation found no remaining issue in that defect class. The
new receiver regression covers direct, extracted, primitive, DOM, parent-window
and unsupported-object calls. The surrounding tests cover stable function
identity, reentry, child errors, popup and frame retirement, and origin checks.

At OfficeIMO code revision `f157343dc3f0be6bc3da71692675e1498d19943c`, the
full single-worker .NET 10 runtime suite passed 646/646 with no skips. Final
.NET 8 focused selections passed 68/68 for document input and auxiliary
windows, and 43/43 for resizable/history state and traversal. These checks use
the same provider commits listed above; they do not constitute a full .NET 8
runtime-suite or Windows/Linux runtime run.

The earlier [combined navigation proof](../../2026-09-22/html-navigation-combined/README.md)
contains the representative rendered workflow at 360 and 720 pixels. This
checkpoint changes a nonvisual script-call boundary; it does not add a new
render. General cross-realm object/property semantics and broader navigation
milestones remain outside this bounded qualification.
