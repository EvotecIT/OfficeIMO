# CSS dimension math and used-value evidence

This evidence qualifies the next retained-provider CSS execution slice at clean source commit
`99c327bd8e84b89153aa8b0b5f4b652ce73c82cb`. OfficeIMO now owns a typed
`<length-percentage>` model for selected sizing and spacing properties, retains unresolved
context through computed-style inspection, and resolves it only at a consuming used-value
boundary.

## Result

`OfficeIMO.Html.Core` exposes provider-independent expression trees for length, percentage,
and mixed length-percentage values. The bounded parser covers literal values and selected
`calc()`, `min()`, `max()`, and `clamp()` arithmetic. The resolver converts absolute units to
canonical CSS pixels and consumes explicit percentage, font, root-font, viewport, and query
container context. Missing context, incompatible types, unsupported units, invalid syntax,
and non-finite results remain distinct outcomes.

The owned property catalog now includes physical `width`, `height`, minimum and maximum
constraints, and physical margin and padding longhands. `HtmlComputedStyle.TryGetTypedValue`
exposes the selected typed value without making layout callers parse provider text. The static
renderer uses the same owned parser and resolver for these properties. Width and horizontal
spacing percentages use the containing width; height constraints use a definite containing
height and remain unresolved when that context is unavailable. Calculated negative sizes are
clamped at the used-value boundary while directly authored invalid negative sizing values are
discarded by property grammar.

The implemented grammar follows [CSS Values and Units](https://drafts.csswg.org/css-values/):
length-percentage values retain a percentage until a reference is available, absolute lengths
resolve to canonical pixels, binary `+` and `-` require CSS whitespace, and unitless zero does
not acquire a length type inside a math expression. The current unit set is `px`, `pt`, `pc`,
`in`, `cm`, `mm`, `q`, `em`, `rem`, the default/small/large/dynamic viewport unit families,
and `cqw`, `cqh`, `cqi`, `cqb`, `cqmin`, and `cqmax`. Font-metric units and wider numeric
functions remain explicit fallback cases.

## Conformance and resource boundaries

The independently authored length corpus contains 28 accepted, rejected, unsupported, and
resolution cases. The shared property corpus contains 35 cases across the complete selected
property slice. End-to-end coverage verifies:

- mixed dimension and percentage typing, canonical forms, absolute conversion, font,
  viewport, and container-relative units;
- arithmetic type checking, precedence, comparison functions, division-by-zero reporting,
  and required binary-operator whitespace;
- tokenizer-compatible decimal spelling and rejection of delimiter unary syntax;
- direct, calculated, and missing-context fallback for horizontal and vertical sizing;
- consistent nonnegative used-value handling in the renderer and query-container prelayout;
- original-source input ceilings before trimming, token, nesting, operation, and argument
  ceilings, plus cooperative cancellation without a partial result.

An independent read-only review found four material defect classes: vertical calculated
percentages using a width reference, whitespace padding bypassing public input ceilings,
malformed numeric and delimiter-unary forms entering the cascade, and calculated negative
query-container sizes disagreeing with final layout. All four were reproduced and corrected.
The single targeted confirmation found no unresolved material issue in those fixes.

## Cross-platform performance gate

The checked reports are [Windows x64](windows-x64.json), [Linux x64](linux-x64.json), and
[macOS Arm64](macos-arm64.json). Each contains 99 validated measurements: three iterations of
11 operations at 10, 100, and 1,000 rows. Every report names the same clean source commit,
reports `trackedSourceDirty: false`, and contains zero budget failures.

The table shows the median large-run time and maximum allocation observed across its three
iterations.

| Platform | Length math | Length-math allocation | Property grammar | Cascade | Traced cascade |
| --- | ---: | ---: | ---: | ---: | ---: |
| Windows x64 | 7.849 ms | 5.07 MiB | 10.777 ms | 897.218 ms | 915.258 ms |
| Linux x64 | 7.090 ms | 5.07 MiB | 10.569 ms | 624.337 ms | 619.890 ms |
| macOS Arm64 | 8.039 ms | 5.07 MiB | 18.274 ms | 1,081.622 ms | 1,155.473 ms |

The new large length-math operation parses and resolves 1,000 mixed pixel, percentage,
font-relative, viewport-relative, and container-relative values. The cascade workload also
uses calculated width and margin declarations through the owned path. Existing operation
budgets were not relaxed; the new operation has its own explicit timing and memory ceilings.

## Compatibility and deployment proof

- The complete `OfficeIMO.Html.Tests` suite passed on Windows with .NET 10 (3,115 tests),
  .NET 8 (3,115), and .NET Framework 4.7.2 (3,114).
- The same source commit passed all 3,115 .NET 10 tests on Linux x64 and macOS Arm64.
- Packed document-only and full `OfficeIMO.Html` consumers passed on .NET Framework,
  .NET 8, and .NET 10. The generated HTML support matrix check also passed.
- The browser bridge package passed its four managed tests. A native Release publish of the
  production Blazor WebAssembly converter on Linux produced 374 files and 106 `.wasm` assets;
  its 3,876,725-byte native runtime contains the required HarfBuzz binding.
- `OfficeIMO.Html.AotSmoke` published and executed as NativeAOT on Linux x64 and macOS Arm64,
  including the typed contextual length contract, SVG, PNG, and searchable-PDF output.

## Dependency decision

No runtime dependency was added or removed. `OfficeIMO.Html.Core` has no package dependency on
either .NET 10 or .NET Framework 4.7.2. The full `OfficeIMO.Html` package still references
AngleSharp 1.7.1 and AngleSharp.Css 1.0.1; its .NET Framework graph also retains the existing
`System.Text.Encoding.CodePages` 8.0.0 transitive dependency and related compatibility
packages.

This remains intentional. The new owned contracts provide useful inspection and rendering
behavior now. The retained CSS provider continues to cover declarations, grouped and nested
rules, and selectors outside the qualified owned subset. Dependency removal remains a later
qualification decision based on shrinking fallback use without reducing real-page behavior.

## Next qualification boundary

The next CSS slice should move selected selector lists, pseudo-classes, namespaces, grouped
rules, or nested qualified rules through the owned contracts according to rendering demand.
Wider Color 4 spaces should follow a concrete rendering consumer. AngleSharp.Css remains
eligible for retirement only when retained-provider fallback is narrow, measured, and safe to
remove. The product backlog remains in
[the HTML roadmap](../../../../../Docs/ROADMAP.md#h3-h6-static-engine-and-dependency-retirement).
