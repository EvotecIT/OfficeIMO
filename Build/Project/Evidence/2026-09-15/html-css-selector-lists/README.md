# CSS selector lists, namespaces, and pseudo-class evidence

This evidence qualifies the retained-provider selector-list slice implemented at
`87f0afa301db2fb7a34eede87e8e2e3a6cf2f276`. The cross-platform performance
reports use clean qualification commit `767508746ec37ff6aee72d2009c40548d93385b1`,
which changes only the large selector workload's retained-heap ceiling after measuring
the Apple Silicon runtime.

## Result

`OfficeIMO.Html.Core` now exposes an immutable `HtmlCssSelectorList`, atomic
`HtmlCssSelectorParser.ParseList` results, stylesheet-scoped
`HtmlCssNamespaceContext`, selector-count and logical-nesting limits, and
cancellation-aware matching. The standalone `Parse` contract remains a
single-complex-selector entry point and classifies a top-level comma as unsupported.

The owned subset now covers:

- selector lists and grouped top-level qualified rules;
- default, prefixed, wildcard, and empty namespace forms for type, universal, and
  attribute selectors;
- `:root`, `:empty`, first/last/only child and of-type forms;
- unfiltered `:nth-child()`, `:nth-last-child()`, `:nth-of-type()`, and
  `:nth-last-of-type()` An+B expressions;
- one basic identifier range in `:lang()`, including inherited `lang` and
  `xml:lang`;
- forgiving `:is()` and `:where()` lists, strict `:not()`, and their Selectors
  Level 4 specificity rules.

Invalid selector lists are discarded atomically. Dynamic state pseudo-classes,
`:has()`, filtered nth selectors, pseudo-elements, nesting, wider `:lang()`
grammar, and selectors outside the qualified subset continue through the retained
provider. For a selector that combines an owned namespace or combinator envelope with
an unsupported pseudo fragment, the full package can delegate only that fragment to
the provider against the current element. The public Core parser never exposes that
provider-assisted representation.

## Conformance and resource boundaries

The independently authored advanced corpus contains 35 parsed, invalid, and
unsupported cases. Twenty-five parsed cases were also evaluated by Chromium
151.0.7922.34; all expected element sets matched. The checked
[differential report](chromium-differential.json) records zero failures.

The automated coverage verifies atomic invalid-list recovery, forgiving and strict
logical-list behavior, namespace prefix declaration order and case sensitivity,
default-namespace subject semantics inside logical pseudo-classes, inherited language
matching, browser-compatible `:empty` behavior, specificity, cancellation, parsing
limits, selector-evaluation limits, and a 5,000-sibling adversarial case. Sibling
positions are cached per parent so structural matching remains linear for the measured
workload.

An independent read-only review reported eight material edge cases. Six were corrected
before its targeted confirmation. That confirmation then exposed two remaining
namespace issues: provider fragments had lost the stylesheet namespace resolver, and
default namespaces were suppressed for every compound inside logical pseudo-classes
instead of only the final subject compound. Both were reproduced and corrected.
Focused tests and the Chromium differential validate the final fixes; no additional
independent review pass was run.

## Cross-platform performance gate

The checked reports are [Windows x64](windows-x64.json),
[Linux x64](linux-x64.json), and [macOS Arm64](macos-arm64.json). Each report
contains 108 validated measurements: three iterations of 12 operations at the small,
normal, and large scales. Every report names clean commit
`767508746ec37ff6aee72d2009c40548d93385b1` and has zero budget failures.

The selector workload uses 100, 1,000, and 6,000 siblings. The table shows median
normal and large elapsed time plus median large allocation and retained managed growth.

| Platform | 1,000 siblings | 6,000 siblings | Elapsed scale | Large allocation | Large retained |
| --- | ---: | ---: | ---: | ---: | ---: |
| Windows x64 | 46.952 ms | 245.173 ms | 5.22x | 47.26 MiB | 4.75 MiB |
| Linux x64 | 22.785 ms | 131.371 ms | 5.77x | 47.49 MiB | 4.75 MiB |
| macOS Arm64, Apple M4 | 21.894 ms | 122.634 ms | 5.60x | 47.27 MiB | 9.24 MiB |

Allocation grows by about 5.7 times for six times as many siblings. The large retained
ceiling is 12 MiB because the macOS runtime retains about twice the managed heap growth
seen on Windows and Linux for this and the existing cascade workloads. The ceiling is
still well below the workload's 80 MiB allocation and managed-peak limits.

## Compatibility and deployment proof

- The complete `OfficeIMO.Html.Tests` suite passed on Windows with .NET 8
  (3,126 tests), .NET 10 (3,126), and .NET Framework 4.7.2 (3,125).
- The same implementation commit passed all 3,126 .NET 10 tests on Linux x64 and
  macOS Arm64.
- Packed document-only and full `OfficeIMO.Html` consumers passed on .NET
  Framework, .NET 8, and .NET 10.
- The browser bridge passed its four managed tests on each Windows target. All 91
  production converter managed tests passed with native linking disabled on Windows.
  A native Release publish on Linux linked the production Blazor WebAssembly
  converter and emitted a 3,876,725-byte `dotnet.native` runtime.
- `OfficeIMO.Html.AotSmoke` published and executed as NativeAOT on Linux x64 and
  macOS Arm64. The Windows host lacks the Visual C++ platform linker, so Windows
  NativeAOT publication was not used as qualification.
- The generated HTML support matrix check, Core and full multi-target Release builds,
  and package-smoke restore/build runs passed.

## Dependency decision

No runtime dependency was added or removed. The checked
[dependency snapshot](dependencies.txt) shows that `OfficeIMO.Html.Core` has no
NuGet dependency on .NET 10. The full `OfficeIMO.Html` package still directly
references AngleSharp 1.7.1 and AngleSharp.Css 1.0.1. Its .NET Framework graph also
retains `System.Text.Encoding.CodePages` 8.0.0 and the existing compatibility
packages.

This remains the useful boundary: OfficeIMO owns the public selector-list and
namespace contracts and executes the qualified rendering path, while AngleSharp.Css
continues to cover wider CSS without reducing existing page support. Dependency
retirement should follow measured fallback use rather than precede working components.

## Next qualification boundary

The next CSS slice should follow rendering demand into nested qualified-rule parsing,
filtered nth selectors, relational or additional pseudo-classes, or a wider property
grammar. Protected-token rewriting and raw-rule reconciliation remain until their
regression fixtures pass entirely through owned parsing. The product backlog remains
in [the HTML roadmap](../../../../../Docs/ROADMAP.md#h3-h6-static-engine-and-dependency-retirement).
