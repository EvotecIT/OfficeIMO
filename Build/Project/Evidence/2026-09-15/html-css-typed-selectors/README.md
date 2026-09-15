# Typed CSS values and selected selector execution evidence

This evidence qualifies the next retained-provider CSS execution slice at clean commit
`4980f6eb105d881871a150beab9118642fd5830c`. OfficeIMO now owns typed constant math,
selected functional colors, and one useful selector and top-level qualified-rule path while
preserving provider fallback for CSS outside the declared subset.

## Result

`OfficeIMO.Html.Core` now parses constant `calc()`, `min()`, `max()`, and `clamp()`
expressions for numbers and percentages. It also exposes typed legacy and modern `rgb()` and
`hsl()` values plus modern `hwb()`. The full package computes clamped opacity and functional
colors from those owned values. Modern unitless HSL and HWB channels use the CSS Color scale,
where `100` corresponds to `100%`.

The owned selector AST and matcher cover type, universal, ID, class, and attribute selectors,
ASCII case modifiers, and descendant, child, adjacent-sibling, and general-sibling
combinators. Matching is cancellation-aware, memoizes visited compound and element states,
and consumes the conversion selector-evaluation budget. Selected top-level qualified rules
retain their owned source and declarations through the managed cascade.

Grouped selectors, namespaces, pseudo-classes, pseudo-elements, nesting, at-rules, and wider
property grammars continue through the retained AngleSharp.Css path. Unsupported selectors
are classified explicitly rather than partially matched. The standards fixtures follow
[CSS Color](https://drafts.csswg.org/css-color/),
[CSS Values and Units](https://drafts.csswg.org/css-values/), and
[Selectors Level 4](https://drafts.csswg.org/selectors-4/). Parsed grammar, cascade acceptance,
computed values, and painted support remain separate claims.

## Conformance and resource boundaries

The independently authored JSON corpora contain 22 typed-value cases and 16 selector cases.
They cover accepted forms, canonical typed results, specificity, matching, invalid syntax,
and valid syntax that requires provider fallback. End-to-end tests also verify:

- exact computed RGBA output for percentage and unitless HSL/HWB forms;
- original-selector retention when declarations require provider fallback;
- CSS binary operator whitespace across comments and iterative unary parsing;
- ASCII-only insensitive comparison and empty-namespace attribute matching;
- cancellation, selector structure limits, memoized adversarial traversal, and conversion
  evaluation-budget exhaustion.

An independent read-only review found six material edge cases across stack safety, selector
traversal bounds, modern and legacy color grammar, ASCII comparison, calculation whitespace,
and namespaced attributes. All six were reproduced and corrected. A targeted confirmation
found no remaining P0-P2 issue in those defect classes.

## Cross-platform performance gate

The checked reports are [Windows x64](windows-x64.json), [Linux x64](linux-x64.json), and
[macOS Arm64](macos-arm64.json). Each report contains 90 validated measurements: three
iterations of ten operations at 10, 100, and 1,000 rows. Every report identifies the same
clean commit and contains zero budget failures.

| Platform | Large typed grammar | Large cascade | Large traced cascade | Cascade allocation | Trace allocation |
| --- | ---: | ---: | ---: | ---: | ---: |
| Windows x64 | 10.541 ms | 760.932 ms | 752.212 ms | 571.07 MiB | 580.81 MiB |
| Linux x64 | 10.672 ms | 797.738 ms | 643.029 ms | 570.88 MiB | 580.73 MiB |
| macOS Arm64 | 21.343 ms | 1,287.397 ms | 1,211.551 ms | 571.17 MiB | 580.89 MiB |

The large cascade validates 3,005 computed-style results. The selector workload adds one
owned scalable qualified rule to the prior corpus and exercises functional color, calculated
opacity, ASCII-insensitive attribute matching, and child combinators. Allocation remains
inside the existing cross-platform ceilings; no budget was relaxed. Timing is retained as
observed evidence rather than normalized into a universal throughput claim.

## Compatibility and deployment proof

- The complete `OfficeIMO.Html.Tests` suite passed on Windows with .NET 8 (3,106 tests),
  .NET 10 (3,106), and .NET Framework 4.7.2 (3,105).
- The same commit passed all 3,106 .NET 10 tests on Linux x64 and macOS Arm64.
- Packed document-only and full `OfficeIMO.Html` consumers passed on .NET Framework,
  .NET 8, and .NET 10.
- All 91 production browser-converter managed tests passed. A native Release WebAssembly
  publish on Linux produced 374 files and 106 `.wasm` assets; its 3,876,725-byte runtime
  contains the required HarfBuzz binding.
- `OfficeIMO.Html.AotSmoke` published and executed as NativeAOT on Linux x64 and macOS Arm64,
  including unitless HSL, calculated opacity, the owned selector path, SVG, PNG, and
  searchable-PDF output.

One initial Windows deadline test and one initial Linux projection-identity test exposed
host-sensitive timing or concurrency behavior outside this CSS slice. Each exact test passed
in isolation, and the complete framework run then passed without a source change.

## Dependency decision

No runtime dependency was added or removed. `OfficeIMO.Html.Core` has no package dependency on
either .NET 10 or .NET Framework 4.7.2. The full `OfficeIMO.Html` package still references
AngleSharp 1.7.1 and AngleSharp.Css 1.0.1; its .NET Framework graph also retains the existing
`System.Text.Encoding.CodePages` 8.0.0 transitive dependency and related compatibility
packages. This is intentional: the new owned contracts are useful now, while unsupported CSS
continues through the qualified provider instead of forcing premature dependency replacement.

## Next qualification boundary

The next CSS work should follow rendering demand: add dimension-aware math and context-bound
percentage resolution, then selected selector lists, pseudo-classes, namespaces, grouped
rules, or nesting. AngleSharp.Css remains eligible for retirement only when owned coverage and
evidence make its fallback narrow enough to remove without reducing real-page behavior. The
product backlog remains in [the HTML roadmap](../../../../../Docs/ROADMAP.md#h3-h6-static-engine-and-dependency-retirement).
