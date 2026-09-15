# Owned CSS property grammar and cascade trace evidence

This evidence qualifies the first retained-provider CSS property slice at clean commit
`43d7bc5b420ed8e581db26059c3dd46ba9028da4`. It records the API boundary, resource
limits, package and deployment shapes, and the same enforced performance workload on
Windows x64, Linux x64, and macOS Arm64.

## Result

OfficeIMO now owns the lossless inline declaration parse, property definitions and typed
values for CSS-wide keywords plus selected `display`, `visibility`, `opacity`, and `color`
grammar. The managed cascade can retain an opt-in, provider-neutral explanation for those
properties. It reports candidate source, selector, layer, specificity, importance, source
order, grammar status, winner decision, inherited/reset behavior, and invalid-at-computed-value
fallback.

Normal style computation does not retain trace graphs. Callers opt in with
`HtmlComputedStyleOptions.IncludeCascadeTraces`. Invalid or cyclic `var()` substitution uses
inheritance for inherited properties and the catalog initial value for other properties.
CSS Color 4 system colors are typed, and the `color` catalog initial value is `CanvasText`.

The default renderer still uses AngleSharp.Css for qualified-rule parsing and selector
matching. This slice removes its inline declaration parser from the managed cascade path; it
does not claim a complete CSS implementation or browser-equivalent rendering. The retained
provider remains replaceable behind OfficeIMO contracts while the owned grammar and execution
surface expands.

## Safety contract

Untrusted conversion applies explicit CSS byte, token, syntax-node, declaration, nesting,
rule, and selector-evaluation limits. Token and syntax-node exhaustion produce separate
`CssTokenLimitExceeded` and `CssSyntaxNodeLimitExceeded` diagnostics with
`MaxCssTokens` and `MaxCssSyntaxNodes` as their limit sources. The trusted prepared-document
overload can remain explicitly unbounded; it no longer inherits the standalone syntax parser's
8 MiB input ceiling.

The property corpus follows the relevant [CSS Values and Units](https://drafts.csswg.org/css-values/),
[CSS Display](https://drafts.csswg.org/css-display-4/),
[CSS Color](https://drafts.csswg.org/css-color/), and
[CSS Cascade](https://drafts.csswg.org/css-cascade-5/) contracts. Parsed grammar support,
cascade acceptance, computed values, and painted output remain separate capability claims.

## Cross-platform performance gate

The checked reports are [Windows x64](windows-x64.json), [Linux x64](linux-x64.json), and
[macOS Arm64](macos-arm64.json). Each contains 90 validated measurements, three iterations of
ten operations at 10, 100, and 1,000 rows. Every report identifies the same clean commit and has
no budget failure.

| Platform | Large property grammar | Large cascade | Large cascade with traces | Trace allocation |
| --- | ---: | ---: | ---: | ---: |
| Windows x64 | 4.573 ms | 995.125 ms | 1,174.259 ms | 526.04 MiB |
| Linux x64 | 4.675 ms | 547.725 ms | 563.952 ms | 525.82 MiB |
| macOS Arm64 | 3.844 ms | 494.230 ms | 492.274 ms | 526.86 MiB |

The large traced workload adds about 4.4 MiB of total allocation over the normal cascade on
each platform. Timing variation is recorded rather than normalized away; all medians remain
inside the platform budget manifest. These measurements cover the declared generated workload
and do not establish universal throughput.

## Compatibility and deployment proof

- The complete `OfficeIMO.Html.Tests` suite passed on Windows with .NET 8 (3,097 tests),
  .NET 10 (3,097), and .NET Framework 4.7.2 (3,096).
- The same clean commit passed all 3,097 .NET 10 tests on Linux x64 and macOS Arm64.
- Packed document-only and full `OfficeIMO.Html` consumers exercised the grammar and trace APIs
  on .NET Framework, .NET 8, and .NET 10.
- The production browser converter passed all 91 managed tests. A native Release WebAssembly
  publish on Linux produced 374 files and 106 `.wasm` assets; the 3,876,725-byte native module
  contains the required HarfBuzz bindings.
- `OfficeIMO.Html.AotSmoke` published and executed as NativeAOT on Linux x64 and macOS Arm64,
  including the owned grammar, cascade trace, SVG, PNG, and searchable-PDF contracts.

The Windows host could compile the managed WASM and NativeAOT inputs but lacked a usable
`emcc` final-link command and the Visual C++ platform linker. Linux supplied the native WASM
proof, and Linux plus macOS supplied actual NativeAOT execution.

## Next qualification boundary

The next CSS slice should add typed color functions and calculated values, then move selected
qualified-rule parsing and selector matching through owned contracts. AngleSharp.Css replacement
should be decided only after those paths preserve the current recovery, cascade, computed-value,
performance, and rendering fixtures. The product backlog remains in
[the HTML roadmap](../../../../../Docs/ROADMAP.md#h3-h6-static-engine-and-dependency-retirement).
