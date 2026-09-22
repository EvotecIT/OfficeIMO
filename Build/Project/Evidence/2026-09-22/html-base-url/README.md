# Dynamic document base URLs

This slice qualifies first-base selection, frozen fallback URLs, DOM insertion,
removal, adoption, cloning and template-content boundaries. AngleSharp owns the
base URL; OfficeIMO uses it for history, navigation, resource loading and capture.
The worker no longer maintains a second base cache or wraps DOM URL getters.

## Reproduction

Initialize the provider submodules at the recorded revisions. From the repository
root, run the real-worker regression classes:

```sh
dotnet test OfficeIMO.Html.Runtime.Tests/OfficeIMO.Html.Runtime.Tests.csproj \
  -f net10.0 \
  --filter 'FullyQualifiedName~RuntimeHistoryTests|FullyQualifiedName~RuntimeBaseUrlTests|FullyQualifiedName~RuntimeBeforeUnloadTests' \
  -- xUnit.MaxParallelThreads=1
```

Repeat with `-f net8.0` for the older runtime. After moving replacement state
onto the owning elements, also run
`ReplacingParserStyleContentRetiresTheOldImportAndCannotRestoreItsSheet` from
`RuntimeScriptLifecycleTests` on both frameworks. The provider's
`DocumentBaseUrlTests`, `MutationVersionTests` and `DOMEventsTests` cover the native
contract. The standalone upstream candidate also runs the full AngleSharp suite.
Exact revisions and results are recorded in `validation.json`, distinguishing
validation before and after the maintainer-requested structural changes.

`RenderProof.cs.txt` is a standalone console fixture referencing `OfficeIMO.Html`,
`OfficeIMO.Html.Runtime` and `OfficeIMO.Html.AngleSharp`. Pass the built runtime
worker DLL and an output directory. It changes the document route, changes a
relative base, changes the route again, fetches data and clicks the relative link.
Both captured states are rendered at 360 and 720 pixels, at a height of 280 pixels.
The four PNG files were visually inspected for the expected content and clipping.

## Contract and provenance

The original regressions derive from the [HTML base-element algorithm](https://html.spec.whatwg.org/multipage/semantics.html#the-base-element)
and selected cases in the web-platform-tests
[`the-base-element` directory](https://github.com/web-platform-tests/wpt/tree/95448d45c5291187f19555a4c6a786ef1ac9885b/html/semantics/document-metadata/the-base-element),
including invalid href reflection and forbidden `data:`/`javascript:` bases.
The WPT revision was recorded during this run; this is selected contract coverage,
not a claim that the complete WPT suite passed.

Independent review found two related template activation cases. Both were
reproduced as failing regressions and corrected: parser-staged inert bases cannot
choose the document base, and replacing template content cannot re-freeze another
active base. Tests also distinguish ordinary template DOM children and content
moved into the document. One full review and one targeted confirmation were used;
the final follow-up fix was validated by the owning agent's regression checks.

Upstream review requested a single `Document` class and no replacement-depth
field on every node. The final implementation keeps frozen-base state in a lazy
internal document helper and replacement state on the template element. The
maintained fork also keeps its stylesheet replacement state on the style element.
Both complete core suites were rerun after these structural changes.

An early concurrent worker run encountered five deadline failures while the host
was running heavy unrelated builds. The serial rerun passed all 215 cases without
relaxing runtime deadlines. A concurrent .NET 8 run was stopped and replaced by
serial validation. Those interrupted/failed runs are not counted as passes.

This evidence covers macOS 27.0 arm64, .NET 10 and .NET 8. It does not qualify
Windows/Linux execution, inherited `about:blank`/`srcdoc` base URLs, CSP base policy,
additional browsing contexts or the full Navigation API. The rendered images are
managed-renderer evidence, not a browser-reference pixel comparison.
