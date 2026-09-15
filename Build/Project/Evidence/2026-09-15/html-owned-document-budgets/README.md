# Owned HTML document and CSS syntax budgets

This evidence qualifies the first OfficeIMO-owned lossless CSS syntax lane and
guards the retained-provider document boundary against time, allocation,
retained-memory, process-memory, output-size and cancellation regressions. It
does not qualify the owned syntax tree as a replacement for the current
selector, cascade or computed-style implementation.

All checked reports use clean commit
`2cc75600111cc4cfe1cc3862d1ef59a8df913902`, .NET 10.0.12, and three isolated
child processes for each operation and scale:

- [Windows x64](windows.json)
- [Linux x64](linux.json)
- [macOS Arm64](macos.json)

Each report contains 63 measurements and zero budget failures. Parse, query,
edit, serialization, conversion and CSS inputs contain 10, 100 and 1,000 report
rows. The cancellation lane starts 10,000-, 25,000- and 100,000-row parses with
a live token and cancels them after the selected provider's parse boundary is entered. Every probe validates a structural
or source fingerprint before recording its result. The committed budget manifest is
`OfficeIMO.Html.Benchmarks/html-owned-document-performance-budgets.json`.

## Normal-scale medians

| Platform | Operation | Elapsed | Allocation | Retained managed heap |
| --- | --- | ---: | ---: | ---: |
| Windows x64 | Parse | 1.431 ms | 1,031.2 KiB | 183.8 KiB |
| Windows x64 | Query | 7.697 ms | 745.3 KiB | 4.1 KiB |
| Windows x64 | Edit | 0.289 ms | 258.4 KiB | 172.1 KiB |
| Windows x64 | Serialize | 7.621 ms | 917.5 KiB | 0 KiB |
| Windows x64 | Conversion with owned document | 1.451 ms | 1,082.3 KiB | 186.5 KiB |
| Windows x64 | Lossless CSS syntax | 3.887 ms | 3,775.0 KiB | 2,220.9 KiB |
| Linux x64 | Parse | 1.429 ms | 1,032.3 KiB | 205.1 KiB |
| Linux x64 | Query | 8.707 ms | 698.8 KiB | 1.5 KiB |
| Linux x64 | Edit | 0.341 ms | 258.1 KiB | 193.5 KiB |
| Linux x64 | Serialize | 8.553 ms | 931.5 KiB | 28.0 KiB |
| Linux x64 | Conversion with owned document | 1.413 ms | 1,083.5 KiB | 207.8 KiB |
| Linux x64 | Lossless CSS syntax | 3.177 ms | 3,774.6 KiB | 2,242.2 KiB |
| macOS Arm64 | Parse | 1.134 ms | 938.7 KiB | 406.9 KiB |
| macOS Arm64 | Query | 4.890 ms | 639.6 KiB | 43.5 KiB |
| macOS Arm64 | Edit | 0.239 ms | 258.3 KiB | 379.4 KiB |
| macOS Arm64 | Serialize | 4.939 ms | 760.3 KiB | 32.5 KiB |
| macOS Arm64 | Conversion with owned document | 1.228 ms | 989.8 KiB | 412.3 KiB |
| macOS Arm64 | Lossless CSS syntax | 2.907 ms | 3,774.8 KiB | 4,100.3 KiB |

The largest retained result is the 1,000-rule CSS tree: 20.94 MiB on Windows,
20.95 MiB on Linux and 38.96 MiB on macOS. The largest measured allocation is 35.06 MiB.
These results pass the 48 MiB retained and 64 MiB allocation ceilings, but they
also identify CSS node/token representation as a later optimization target.

## Provider decision delta

The current Windows provider comparison is retained in
[provider-windows.json](provider-windows.json). It uses the same generated
100-row document and 105-rule stylesheet as the original provider baseline.

| Scenario | Median elapsed | Median allocation | Retained heap per result |
| --- | ---: | ---: | ---: |
| AngleSharp HTML parse | 0.705 ms | 388.4 KiB | 270.1 KiB |
| OfficeIMO owned document | 1.313 ms | 1,033.7 KiB | 204.3 KiB |
| Conversion with native document | 0.833 ms | 441.8 KiB | 273.1 KiB |
| Conversion with owned document | 1.714 ms | 1,086.9 KiB | 207.0 KiB |
| AngleSharp.Css syntax | 9.146 ms | 1,939.7 KiB | 691.0 KiB |
| OfficeIMO lossless CSS syntax | 3.905 ms | 3,775.1 KiB | 2,241.2 KiB |
| OfficeIMO cascade for 305 elements | 228.247 ms | 50,398.4 KiB | 1,079.8 KiB |

Weak provider projections reduce the directly comparable Windows owned-document
retention from 638.9 KiB to 204.3 KiB and conversion native-plus-owned retention
from 641.9 KiB to 207.0 KiB, both about 68 percent. A caller that still holds a
native document keeps that document identity; OfficeIMO can rebuild its node
maps around the same document without making the provider graph permanent.

The CSS syntax timings measure different result contracts. The OfficeIMO result
retains exact source, trivia, unknown syntax, nested component values, recovery
nodes and source spans, which accounts for its larger graph. AngleSharp.Css
continues to supply the qualified selector, cascade and computed-style path.

## Reproduction and compatibility checks

Run the enforced budget gate from the repository root:

```powershell
dotnet run -c Release -f net10.0 --project ./OfficeIMO.Html.Benchmarks -- `
  --owned-document-verify-budgets --repeat 3 --json owned-document.json
```

The same candidate also passed:

- 3,071 `OfficeIMO.Html.Tests` tests on Windows x64 and Linux x64 with .NET 10;
- locally packed document-only and full consumers on .NET Framework 4.7.2,
  .NET 8 and .NET 10;
- the focused CSS and projection-lifetime contracts on macOS Arm64;
- actual `OfficeIMO.Html.AotSmoke` NativeAOT publish and execution on Linux x64
  and macOS Arm64;
- all 91 production browser-converter tests and a native Release WebAssembly
  publish containing 369 files, 106 `.wasm` assets and the required HarfBuzz
  symbols.

The source corpus has fourteen compact cases covering duplicate and unknown
declarations, unknown and nested at-rules, custom-property blocks, strings and
URLs containing delimiters, escapes, CDO/CDC handling, unclosed rules and
invalid declaration recovery, block-valued unknown declarations and mismatched
delimiter recovery. This is the first owned syntax qualification
slice. Broader selected standards fixtures, property grammar and selector and
cascade integration remain required before removing AngleSharp.Css.
