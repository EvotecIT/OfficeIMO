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
runtime `RuntimeDocumentOpenTests` check that the parent adopts the entry URL,
loses its old `<base>` element, retains its origin, and uses its own fallback
base rather than the child's inherited base. The runtime cases execute from
child-owned timer tasks; dispatching a child event from the parent would retain
the parent as the entry global and test a different contract. These assertions
follow the [HTML document-open steps](https://html.spec.whatwg.org/multipage/dynamic-markup-insertion.html#document-open)
and [fallback-base rules](https://html.spec.whatwg.org/multipage/urls-and-fetching.html#document-base-urls).
Chromium 151 differed in a separate browser observation: the blank case kept
the parent's former `/assets/` base, while the srcdoc case kept its former HTTP
URL. This is an interop limit, not evidence that either behavior is universal.
