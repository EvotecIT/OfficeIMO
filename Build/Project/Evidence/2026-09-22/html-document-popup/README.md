# Document input and surviving auxiliary-window evidence

The combined workflow restores shared resizable-buffer views and replaces an
iframe through `document.open/write/close`. Two writes preserve the existing
paragraph object, including text split at a character reference. The popup stays
open across root reload: its counter advances from one to two, and its saved
opener reference accesses and messages the replacement root document.

The four PNGs show the checked state before and after reload at 360 and 720
pixels. Runtime assertions also check the frame URL, document-removal mutation
record, history buffer aliases, view lengths, popup base, message origin and
message source. These are bounded contracts, not a complete browser or
navigation qualification.

`RenderProof.cs.txt` is executable console source. Reference OfficeIMO.Html,
OfficeIMO.Html.Runtime and OfficeIMO.Html.AngleSharp, then pass the built worker
DLL and an output directory. `validation.json` records source revisions, selected
checks, review coverage, artifact hashes and remaining qualification limits.

The later reverse-call qualification covers both `about:blank` and `about:srcdoc`
children opening an active HTTP parent. Native `DocumentOpenLifecycleTests` and
runtime `RuntimeDocumentOpenTests` check that the parent adopts the entry URL
in both `document.URL` and `location.href`, loses its old `<base>` element,
retains its origin, and uses its own fallback
base rather than the child's inherited base. The runtime cases execute from
child-owned timer tasks; dispatching a child event from the parent would retain
the parent as the entry global and test a different contract. These assertions
follow the [HTML document-open steps](https://html.spec.whatwg.org/multipage/dynamic-markup-insertion.html#document-open)
and [fallback-base rules](https://html.spec.whatwg.org/multipage/urls-and-fetching.html#document-base-urls).
Chromium 151 differed in a separate browser observation: the blank case kept
the parent's former `/assets/` base, while the srcdoc case kept its former HTTP
URL. This is an interop limit, not evidence that either behavior is universal.

The [unfinished-token probe](unfinished-token-performance.json) checks 2,000 to
16,000 one-character writes inside a still-open start-tag attribute. At 16,000
writes, the Release-mode median fell from 467 ms to 6.9 ms; a quoted `>` before
the writes took 7.2 ms after the fix. The probe verifies the completed attribute
length, and the native and runtime regression suites cover token boundaries.
These measurements are local to macOS arm64 and do not qualify other unfinished
token types.

To repeat the probe, build the selected AngleSharp commit in Release for
`net10.0`, create a small `net10.0` console project referencing that assembly,
and use `UnfinishedTokenProbe.cs.txt` as `Program.cs`. Run once normally and
once with `QUOTED_BRACKET=true`; the latter puts `>` inside the quoted value
before the measured writes. The probe discards its first sample per size and
checks the completed attribute length after each timed loop.

The input-stream URL checkpoint uses AngleSharp fork commit `549a94650` and
OfficeIMO code commit `c68144314`. It checks implicit `write()` and `writeln()`
opens from a same-origin parent into a completed child, and a child-owned timer
writing into its parent. The receiver takes the entry document URL, strips its
fragment when the documents differ, and updates `location.href` without adding
history. Runtime tests also check one- and two-entry history state and reject
cross-origin frame writes before erasing content. The native lifecycle suite
passes 19/19; its full .NET 10 suite passes 4,206 with one skip. The runtime
document-input class passes 21/21 on both .NET 10 and .NET 8; adjacent .NET 10
script, history, state and frame-base checks pass 84/84. The native
`netstandard2.0` Release build has no warnings or errors. An independent
read-only review found no actionable issue in this contract.

These are selected adaptations of the URL and history behavior in WPT's
[opening-the-input-stream tests](https://github.com/web-platform-tests/wpt/tree/02c25cc3a34a70a6225147f76cec312483df4706/html/webappapis/dynamic-markup-insertion/opening-the-input-stream),
not an official WPT harness run. In particular, the
[`url-entry-document-sync-call.window.js` case](https://github.com/web-platform-tests/wpt/blob/02c25cc3a34a70a6225147f76cec312483df4706/html/webappapis/dynamic-markup-insertion/opening-the-input-stream/url-entry-document-sync-call.window.js)
invokes a child-defined `window` function from the parent. The runtime does
not yet forward arbitrary child window functions across realms, so the direct
cross-realm document-method tests do not qualify that case. The WPT history
case also uses a popup; the current one- and two-entry checks run on the root
browsing context.
