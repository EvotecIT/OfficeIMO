# Local frame base URL evidence

This fixture keeps a srcdoc report usable after its parent changes its base URL
and history entry. The child fetches `/assets/data.txt`, retains `about:srcdoc`
as its document URL and keeps the original `/assets/` fallback base. The four PNGs
show the loaded and changed states at 360 and 720 pixels, inspected on macOS arm64.

`validation.json` records the exact source revisions, commands and test results.
`RenderProof.cs.txt` is the executable console source: reference OfficeIMO.Html,
OfficeIMO.Html.Runtime and OfficeIMO.Html.AngleSharp, then pass the built worker
DLL path and an output directory. Build with the SDK recorded in the manifest.
`source.html` and `child.html` are the authored input, not captured browser output.

The native and worker regressions cover blank and srcdoc identity, relative
classic/module/fetch URLs, frozen creator bases, child base changes, inherited
origin independent of resource bases, sandboxed ancestors, nested frame realms,
detached containers/fragments and suppression of recursive HTTP embedding.

Acceptance derives from the [HTML fallback-base rules](https://html.spec.whatwg.org/multipage/urls-and-fetching.html#fallback-base-url)
and selected base-URL WPT sources listed in the manifest. These are original
regressions, not a full WPT conformance run. Popup initiators, document.open URL
rewriting, replacement child navigation and other operating systems remain
outside this evidence. The independent review's two reproduced findings were
fixed before the final validation; its targeted confirmation is recorded in the manifest.
