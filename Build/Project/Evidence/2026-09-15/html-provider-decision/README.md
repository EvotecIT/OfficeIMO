# HTML retained-provider decision baseline

This evidence separates the current HTML parser, CSS syntax/cascade, and scripted-runtime DOM decisions. It supports a retained-provider OfficeIMO.Html release candidate; it does not claim that the default package graph is dependency-free.

## Decision

| Area | Current decision | Revisit when |
| --- | --- | --- |
| AngleSharp HTML parser | Retain behind `HtmlDocumentEngine` | A managed parser passes the selected recovery, contextual-fragment, encoding, bounds, differential, and performance gates |
| AngleSharp.Css | Retain for current behavior and build the owned lossless CSS syntax lane first | The owned path preserves supported syntax and recovery and produces the same selector, cascade-trace, and computed-value contracts without protected-token or raw-rule recovery |
| Retained AngleSharp runtime DOM fork | Retain for the trusted managed runtime and govern it separately | Upstreamable hooks or an owned live-DOM replacement pass the runtime conformance and lifecycle gates |

The parser decision is based on behavior and cost rather than package count alone. The native parser is fast on the representative workload. Most of the current parse-path increase comes from retaining and projecting both the provider and OfficeIMO-owned graphs. Replacing the parser would not by itself remove AngleSharp.Css or the separately compiled runtime DOM fork.

## Reproduction

Build and run the process-isolated provider lanes:

```powershell
dotnet build OfficeIMO.Html.Benchmarks/OfficeIMO.Html.Benchmarks.csproj -c Release -f net10.0
dotnet run --project OfficeIMO.Html.Benchmarks/OfficeIMO.Html.Benchmarks.csproj -c Release -f net10.0 --no-build -- --provider-evidence --repeat 3 --json provider.json
```

Every scenario runs in a fresh child process after three warm operations. Reports record elapsed time, total allocation, retained managed heap per result from a retained batch of eight, absolute process peak, result validation, assembly versions, runtime, operating system, architecture, commit, and tracked-source state. The four HTML lanes must return the same exact 420-element structural fingerprint; the command fails instead of publishing incomparable results.

The HTML lanes use the same generated 100-row report. The CSS syntax lane parses the generated stylesheet only; the OfficeIMO cascade lane parses the styled document once and then computes styles for 305 elements. Their elapsed times must not be divided into a parser performance ratio because the cascade lane performs substantially more work.

The checked reports are [Windows](windows.json), [Linux](linux.json), and [macOS](macos.json). Each report must identify the same clean commit before it is used as cross-platform decision evidence.

## Package and source snapshot

The .NET 10 package graphs at this baseline are:

| Package | Direct third-party runtime packages |
| --- | --- |
| `OfficeIMO.Html.Core` | None |
| `OfficeIMO.Html.AngleSharp` | AngleSharp 1.7.1; System.Text.Encoding.CodePages 8.0.0 |
| `OfficeIMO.Html` | AngleSharp 1.7.1; AngleSharp.Css 1.0.1; the two OfficeIMO packages above and shared OfficeIMO owners |

The document-only packed consumer installs `OfficeIMO.Html.AngleSharp` and exercises parse, query, contextual fragment, import, edit, and serialization without referencing the conversion, graphics, or PDF layers. The full packed consumer exercises the owned document API, Markdown, rendering, PDF, and Office adapters. Both run on .NET Framework 4.7.2, .NET 8, and .NET 10.

The source-maintenance snapshot contains 7 C# files and 984 lines in the static AngleSharp adapter, 15 files and 1,089 lines in `OfficeIMO.Html.Core`, and 720 files and 84,402 lines in the retained runtime DOM fork. Nine CSS style-engine files directly reference AngleSharp DOM or CSS contracts. These counts describe maintenance surface; they are not code-quality or feature-completeness scores.

At the dated baseline, the current upstream stable releases were AngleSharp 1.8.1 and AngleSharp.Css 1.0.1. OfficeIMO remains pinned to AngleSharp 1.7.1 because the retained runtime fork is independently based on that version. An upgrade is a qualified dependency-maintenance change, separate from the keep-or-replace decision.

## Qualification boundaries

- Parsing is inert: it does not execute scripts or fetch resources.
- Cancellation inside the retained parser is cooperative. Source limits run before parsing, while node and depth limits run after the provider has built its native tree.
- The managed static renderer is browser-free. The current package graph is not third-party dependency-free.
- NativeAOT executes the public document, contextual-fragment, SVG, PNG, and searchable-PDF smoke on Linux x64. The production Blazor WebAssembly converter links the same HTML provider and rendering graph with native WebAssembly assets; browser-specific feature claims remain limited to the generated support matrix.
- The trusted scripted runtime and an untrusted-content sandbox are different profiles. Retaining a worker process does not establish operating-system isolation for hostile scripts.
