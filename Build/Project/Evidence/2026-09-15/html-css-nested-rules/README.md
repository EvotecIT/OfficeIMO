# CSS nested qualified-rule evidence

This evidence qualifies the nested qualified-rule slice implemented at clean commit
`acbf2c7201ee8a7b0fb5cb058799fc86ab87a544`.

## Result

`OfficeIMO.Html.Core` now exposes immutable child declaration and rule snapshots on
`HtmlCssRule`. The full renderer pairs those owned nodes with the provider stylesheet
and resolves nested qualified rules before cascade evaluation.

The qualified subset covers:

- nesting selectors that contain `&` and implicit descendant nesting;
- selector-list cross products, including escaped and commented commas;
- declarations before, between, and after nested rules in authored order;
- nested qualified rules inside supported conditional grouping contexts;
- namespace-aware owned selectors inside nested envelopes;
- hybrid selectors whose unsupported pseudo fragment still uses the retained provider;
- invalid nested selector-list isolation and existing CSS recovery behavior.

The renderer checks the resolved selector count and total selector characters before
materializing a nesting cross product. `MaxCssSelectorsPerRule` defaults to 256 and
`MaxCssSelectorCharacters` defaults to 64 KiB. A rejected expansion reports
`CssSelectorExpansionLimitExceeded` instead of allocating the full product.

## Review and conformance

The independently authored nested-rule corpus covers nesting selectors, implicit
descendants, selector lists, namespaces, conditional groups, declarations around
nested rules, invalid recovery, hybrid provider fallback, specificity, and resource
limits.

One independent read-only review found four material edge cases: exponential selector
expansion, ordering around provider-only rules, namespace fallback, and commas hidden
by escapes or comments. All four were reproduced and fixed. The permitted targeted
confirmation closed those findings and exposed one additional hybrid-selector
specificity regression involving `:has(#id)`. That regression was also reproduced and
fixed. Focused and full-suite tests exercise the final behavior; no further independent
review pass was run.

## Cross-platform performance gate

The checked reports are [Windows x64](windows-x64.json),
[Linux x64](linux-x64.json), and [macOS Arm64](macos-arm64.json). Each report contains
three validated measurements at each scale, names clean implementation commit
`acbf2c7201ee8a7b0fb5cb058799fc86ab87a544`, and records zero budget failures.

The table shows median elapsed time, allocation, retained managed growth, and managed
peak for the large nested-rule workload.

| Platform | Elapsed | Allocation | Retained | Managed peak |
| --- | ---: | ---: | ---: | ---: |
| Windows x64 | 445.678 ms | 344.88 MiB | 0.055 MiB | 99.14 MiB |
| Linux x64 | 445.813 ms | 347.93 MiB | 0.008 MiB | 103.78 MiB |
| macOS Arm64, Apple M4 | 361.602 ms | 344.20 MiB | 0.016 MiB | 136.91 MiB |

The benchmark deliberately exercises selector-list cross products and mixed declaration
runs at increasing scale. Its fingerprint and result count are validated before a
measurement can satisfy the budget gate.

## Compatibility and deployment proof

- The complete `OfficeIMO.Html.Tests` suite passed on Windows with .NET 8 (3,138
  tests), .NET 10 (3,138), and .NET Framework 4.7.2 (3,137).
- The same implementation commit passed all 3,138 .NET 10 tests on Linux x64 and
  macOS Arm64.
- The 14 focused nested-rule tests passed after the final specificity correction.
- Packed document-only and full `OfficeIMO.Html` consumers passed on .NET Framework,
  .NET 8, and .NET 10.
- The browser bridge passed its four managed tests on each Windows target. All 91
  production converter managed tests passed on Windows.
- The production Blazor WebAssembly converter completed a native Release publish on
  Linux. Its 3,876,725-byte `dotnet.native` runtime contains the required HarfBuzz
  symbol; the exact output is recorded in [wasm-publish.txt](wasm-publish.txt).
- `OfficeIMO.Html.AotSmoke` published and executed as NativeAOT on Linux x64 and
  macOS Arm64.
- The generated HTML support matrix check and package-smoke restore/build runs passed.

## Dependency decision

No runtime dependency was added or removed. The checked
[dependency snapshot](dependencies.txt) shows that `OfficeIMO.Html.Core` has no NuGet
dependency on .NET 10. The full `OfficeIMO.Html` package continues to reference
AngleSharp 1.7.1 and AngleSharp.Css 1.0.1 directly. Its .NET Framework graph continues
to carry `System.Text.Encoding.CodePages` 8.0.0 and the existing compatibility
packages.

OfficeIMO now owns the nested-rule contract and qualified rendering path. The retained
provider continues to preserve wider CSS support inside unsupported selector fragments,
unknown rules, and grammar outside the qualified subset. This keeps the component
useful today while preserving the boundary needed to reduce dependencies later.

## Next qualification boundary

The next CSS work should follow measured rendering demand into filtered nth selectors,
relational pseudo-classes beyond the current hybrid path, conditional evaluation, or
another property/layout slice. Visual acceptance thresholds for the advanced held-out
corpus remain the next H4 promotion gate. Parser replacement starts only when the
selected parser corpus or a blocked product behavior gives a measured reason to retire
the retained HTML parser.

The open work remains in [the HTML roadmap](../../../../../Docs/ROADMAP.md#h3-h6-static-engine-and-dependency-retirement).
