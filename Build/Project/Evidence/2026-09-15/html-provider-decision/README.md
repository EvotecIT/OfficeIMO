# HTML retained-provider decision baseline

This evidence separates the current HTML parser, CSS syntax/cascade, and scripted-runtime DOM decisions. It supports a retained-provider OfficeIMO.Html release candidate; it does not claim that the default package graph is dependency-free.

## Decision

| Area | Current decision | Revisit when |
| --- | --- | --- |
| AngleSharp HTML parser | Retain behind `HtmlDocumentEngine` | A managed parser passes the selected recovery, contextual-fragment, encoding, bounds, differential, and performance gates |
| AngleSharp.Css | Retain for current behavior and build the owned lossless CSS syntax lane first | The owned path preserves supported syntax and recovery and produces the same selector, cascade-trace, and computed-value contracts without protected-token or raw-rule recovery |
| Retained AngleSharp runtime DOM fork | Retain for the trusted managed runtime and govern it separately | Upstreamable hooks or an owned live-DOM replacement pass the runtime conformance and lifecycle gates |

The parser decision is based on behavior and cost rather than package count alone. The native parser is fast on the representative workload. Most of the current parse-path increase comes from retaining and projecting both the provider and OfficeIMO-owned graphs. Replacing the parser would not by itself remove AngleSharp.Css or the separately compiled runtime DOM fork.

The subsequent [owned property grammar and cascade trace evidence](../html-css-property-cascade/README.md)
qualifies the first selected property slice, moves inline declarations through the owned syntax
tree, and retains AngleSharp.Css for qualified rules and selector matching.

## Reproduction

Build and run the process-isolated provider lanes:

```powershell
dotnet build OfficeIMO.Html.Benchmarks/OfficeIMO.Html.Benchmarks.csproj -c Release -f net10.0
dotnet run --project OfficeIMO.Html.Benchmarks/OfficeIMO.Html.Benchmarks.csproj -c Release -f net10.0 --no-build -- --provider-evidence --repeat 3 --json provider.json
```

Every scenario runs in a fresh child process after three warm operations. Reports record elapsed time, total allocation, retained managed heap per result from a retained batch of eight, process peak where the runtime exposes it, result validation, assembly versions, runtime, operating system, architecture, commit, and tracked-source state. The four HTML lanes must return the same exact 420-element structural fingerprint; the command fails instead of publishing incomparable results.

The HTML lanes use the same generated 100-row report. The CSS syntax lane parses the generated stylesheet only; the OfficeIMO cascade lane parses the styled document once and then computes styles for 305 elements. Their elapsed times must not be divided into a parser performance ratio because the cascade lane performs substantially more work.

The checked reports are [Windows](windows.json), [Linux](linux.json), and [macOS](macos.json). Each report must identify the same clean commit before it is used as cross-platform decision evidence.

All reports use clean commit `74d80d8ea32a65836bd35543264ac0a6620b7b66`, .NET 10.0.12, three measured child processes per scenario, and HTML fingerprint `85178b304a9c260dd560cba117fea0d59e828b59547f7adc34f1a0f942e88f0c`.

### HTML parser and projection

| Platform | Scenario | Median elapsed | Median allocation | Retained heap per result |
| --- | --- | ---: | ---: | ---: |
| Windows x64 | Native AngleSharp parse | 0.639 ms | 388.4 KiB | 270.1 KiB |
| Windows x64 | Owned `HtmlDocument` parse and projection | 1.251 ms | 1,033.6 KiB | 638.9 KiB |
| Windows x64 | Conversion document, retained native graph | 1.065 ms | 441.8 KiB | 273.1 KiB |
| Windows x64 | Conversion document, retained native and owned graphs | 1.397 ms | 1,086.9 KiB | 641.9 KiB |
| Linux x64 | Native AngleSharp parse | 0.765 ms | 388.1 KiB | 270.1 KiB |
| Linux x64 | Owned `HtmlDocument` parse and projection | 1.306 ms | 1,031.7 KiB | 641.0 KiB |
| Linux x64 | Conversion document, retained native graph | 0.961 ms | 441.5 KiB | 273.2 KiB |
| Linux x64 | Conversion document, retained native and owned graphs | 1.462 ms | 1,084.9 KiB | 644.0 KiB |
| macOS Arm64 | Native AngleSharp parse | 0.643 ms | 388.1 KiB | 545.4 KiB |
| macOS Arm64 | Owned `HtmlDocument` parse and projection | 1.012 ms | 860.4 KiB | 1,171.5 KiB |
| macOS Arm64 | Conversion document, retained native graph | 0.720 ms | 441.5 KiB | 551.5 KiB |
| macOS Arm64 | Conversion document, retained native and owned graphs | 1.102 ms | 913.6 KiB | 1,167.9 KiB |

The retained native graph is the smaller and faster result in this workload. Creating the provider-neutral owned graph costs roughly another 0.4-0.7 ms and 472-661 KiB of allocation. This makes dual-graph lifetime and projection the first optimization target. It does not justify starting an HTML parser rewrite while the provider remains fast and behaviorally qualified.

### CSS syntax and cascade

| Platform | Scenario | Median elapsed | Median allocation | Retained heap per result |
| --- | --- | ---: | ---: | ---: |
| Windows x64 | AngleSharp.Css syntax, 105 rules | 9.747 ms | 1,939.7 KiB | 691.0 KiB |
| Windows x64 | OfficeIMO cascade, 305 elements | 248.175 ms | 49,710.9 KiB | 1,079.9 KiB |
| Linux x64 | AngleSharp.Css syntax, 105 rules | 7.888 ms | 1,939.4 KiB | 691.0 KiB |
| Linux x64 | OfficeIMO cascade, 305 elements | 221.814 ms | 49,710.9 KiB | 1,079.9 KiB |
| macOS Arm64 | AngleSharp.Css syntax, 105 rules | 6.938 ms | 1,939.4 KiB | 1,387.7 KiB |
| macOS Arm64 | OfficeIMO cascade, 305 elements | 203.720 ms | 49,710.9 KiB | 2,160.8 KiB |

The CSS rows measure different operations. They show that raw syntax parsing is a small part of the current styled-document cost; selector matching, cascade, computed values and OfficeIMO recovery paths remain the larger optimization surface. `Process.PeakWorkingSet64` returned zero for the isolated macOS provider probes, so those raw fields are unavailable rather than a zero-memory observation. Retained managed heap remains recorded on every platform, and the separately supervised H4 process reports macOS peak working set.

## Package and source snapshot

The locally packed 3.4.4 candidate package graphs at this baseline are:

| Package | Candidate `.nupkg` size | Direct third-party runtime packages |
| --- | ---: | --- |
| `OfficeIMO.Html.Core` | 99,451 B | None |
| `OfficeIMO.Html.AngleSharp` | 89,464 B | AngleSharp 1.7.1; System.Text.Encoding.CodePages 8.0.0 |
| `OfficeIMO.Html` | 2,870,426 B | AngleSharp 1.7.1; AngleSharp.Css 1.0.1; the two OfficeIMO packages above and shared OfficeIMO owners |

The document-only packed consumer installs `OfficeIMO.Html.AngleSharp` and exercises parse, query, contextual fragment, import, edit, and serialization without referencing the conversion, graphics, or PDF layers. The full packed consumer exercises the owned document API, Markdown, rendering, PDF, and Office adapters. Both run on .NET Framework 4.7.2, .NET 8, and .NET 10.

The source-maintenance snapshot contains 7 C# files and 1,019 lines in the static AngleSharp adapter, 15 files and 1,089 lines in `OfficeIMO.Html.Core`, and 720 files and 84,402 lines in the retained runtime DOM fork. Nine CSS style-engine files directly reference AngleSharp DOM or CSS contracts. These counts describe maintenance surface; they are not code-quality or feature-completeness scores.

At the dated baseline, the current upstream stable releases were AngleSharp 1.8.1 and AngleSharp.Css 1.0.1. OfficeIMO remains pinned to AngleSharp 1.7.1 because the retained runtime fork is independently based on that version. An upgrade is a qualified dependency-maintenance change, separate from the keep-or-replace decision.

## Qualification boundaries

- Parsing is inert: it does not execute scripts or fetch resources.
- Cancellation inside the retained parser is cooperative. Source limits run before parsing, while node and depth limits run after the provider has built its native tree.
- The managed static renderer is browser-free. The current package graph is not third-party dependency-free.
- NativeAOT executes the public document, contextual-fragment, SVG, PNG, and searchable-PDF smoke on Linux x64 and macOS Arm64. The production Blazor WebAssembly converter links the same HTML provider and rendering graph with native WebAssembly assets; the checked publish contained 374 files, 106 `.wasm` assets and 97,241,027 bytes. Browser-specific feature claims remain limited to the generated support matrix.
- The trusted scripted runtime and an untrusted-content sandbox are different profiles. Retaining a worker process does not establish operating-system isolation for hostile scripts.
