# OfficeIMO RTF support matrix

Last reviewed: 2026-09-30

Use this matrix to check whether an RTF workflow is fully modeled, broadly supported with a named fidelity boundary, preserved without semantic editing, or available only for extraction. The [machine-readable capability catalog](officeimo.rtf-capabilities.json) exposes the same classifications to tooling.

Status meanings:

- `Full`: the stated contract has a public API and focused evidence.
- `Broad`: useful coverage exists with a named fidelity boundary.
- `Preserved`: source syntax survives losslessly but is not fully semantic.
- `Extractive`: useful content and provenance are recovered without round-trip claims.

## Packages and conversion graph

```text
                        trusted round trip
                    HTML <--------------> RTF
                      ^                    |
                      |                    | semantic + result report
                 web-safe HTML             v
Reader chunks <------ RTF <-------------> Word/DOCX
                      |
                      +------> Markdown
                      |
                      +------> PDF (visual export)
                      ^
                      +------- PDF (extractive import)
```

| Area | Owner | Current contract |
| --- | --- | --- |
| Core syntax, semantic model, writer, limits, reports, and editing | `OfficeIMO.Rtf` | Reusable RTF engine. |
| Word mapping and document workflows | `OfficeIMO.Word.Rtf` | Thin bridge over `OfficeIMO.Word`. |
| Web-safe and round-trip HTML | `OfficeIMO.Html.Rtf` | URL/resource policy plus semantic HTML mapping over the independent HTML and RTF engines. |
| Markdown | `OfficeIMO.Rtf.Markdown` | Semantic `RtfDocument`/`MarkdownDoc` bridge. |
| PDF | `OfficeIMO.Rtf.Pdf` | Visual export and extractive import through `OfficeIMO.Pdf`. |
| Reader | `OfficeIMO.Reader.Rtf` | Bounded chunk, table, visual, warning, and provenance extraction. |

The packages target `netstandard2.0`, .NET 8, and .NET 10; Windows builds also target .NET Framework 4.7.2.

## P0: safe and diagnosable ingestion

| Capability | Status | Contract and evidence |
| --- | --- | --- |
| <!-- capability:core-parse --> Core parsing | Full | `RtfDocument.Read/Load` parse groups, controls, text, binary payloads, and recoverable malformed input. |
| <!-- capability:lossless-roundtrip --> Lossless round trip | Full | `ToRtfLossless` and lossless save preserve source bytes, unknown destinations, binary payloads, and trailing bytes. Semantic writes are normalized. |
| <!-- capability:bounded-ingestion --> Bounded ingestion | Full | `RtfReadOptions.CreateUntrustedProfile()` caps input bytes/chars, depth, tokens, groups, text, binary, images, objects, and semantic blocks. |
| <!-- capability:cancellation --> Cancellation | Full | String, byte, file, and stream routes support cooperative cancellation through tokenization and semantic binding. |
| <!-- capability:unknown-diagnostics --> Unknown and preserve-only destinations | Full | Stable diagnostics identify unknown ignorable destinations and classified advanced families while the lossless path retains syntax. |
| <!-- capability:shared-report --> Shared conversion truth | Full | `RtfConversionReport` is used across semantic adapters. `RequireNoLoss()` rejects flattened, omitted, blocked, or error diagnostics. `ToRtfResult()` reports normalization loss, including unbound syntax and stale encapsulated HTML. PDF uses `PdfConversionReport` with `RtfAction` details. |
| <!-- capability:web-safe-html --> HTML output profiles | Full | `CreateWebSafeProfile()` blocks unsafe URL schemes, private metadata, and inline payloads by default. `CreateRoundTripProfile()` is explicit trusted output. |

The untrusted profile is intentionally conservative. Applications can clone or create `RtfReadOptions` with different limits, but uploaded RTF should not use the compatibility-oriented default without a host boundary.

## P1: interoperability and adapters

| Capability | Status | Contract and boundary |
| --- | --- | --- |
| <!-- capability:dbcs --> DBCS and font charset switching | Broad | Windows code pages 932, 936, 949, and 950 are decoded through `System.Text.Encoding.CodePages`; Word-style `fcharset` changes are honored. Composite-font parity is not claimed. |
| <!-- capability:outlook-html --> Outlook HTML encapsulation | Broad | `fromhtml`, `htmltag`, `htmlrtf`, and `mhtmltag` are recognized, modeled, written, and preferred by HTML conversion. Focused grammar tests are backed by a reproducible Outlook 16 in-memory MAPI `RTFBody` fixture with genuine `fromhtml` and `htmltag` controls. |
| <!-- capability:styles-numbering --> Styles and numbering | Broad | Effective paragraph and character formatting resolves inheritance, hidden text, explicit resets, and document defaults. Style-gallery hiding is separate from hidden text. Semantic writing materializes supported effective body controls for native readers by default and reports the lost inheritance distinction; `MaterializeStyleFormatting = false` retains editable inheritance. Common list definitions, override starts, shared counters, and level restarts drive Word, HTML, Markdown, PDF, and Reader output. Full character effects, effective table styles, picture bullets, unsupported number formats, themes, and latent styles remain gaps. |
| <!-- capability:nested-tables --> Nested tables | Broad | Native core, HTML, and Word nesting is supported. Markdown, PDF, and Reader flatten nested tables and report that action. |
| <!-- capability:images --> Images | Broad | PNG, JPEG, and supported DIB data use the shared drawing layer. Picture scale and crop geometry map through Word, HTML, and PDF, with validation and fallback diagnostics. Markdown has a media callback; PDF accepts a WMF/EMF converter callback. |
| <!-- capability:notes-markdown --> Markdown notes and media | Broad | Footnotes/endnotes become Markdown references and definitions. Headers/footers remain explicit diagnostic omissions. |
| <!-- capability:advanced-destinations --> Advanced RTF families | Preserved | Move revisions, protection exceptions, index/TOC entries, custom XML, smart tags, and related destinations are classified and diagnosed. Not every family is semantically editable. |
| <!-- capability:word-workflows --> Word workflows | Broad | Result-bearing mail merge, cross-run find/replace, field update, append/merge, and comparison route through `OfficeIMO.Word` and return combined conversion/workflow reports. |
| <!-- capability:pdf-adapter --> PDF bridge | Broad | Export maps semantic layout, tables, images, links, notes, section-owned running stories, and equal-width section columns. Page-starting section geometry, page-number continuation/restarts, and parity padding are supported. Unequal columns, column-starting section transitions, and continuous-section page-setting changes produce flattening diagnostics. Import is logical extraction, not lossless PDF reconstruction. |
| <!-- capability:reader-adapter --> Reader bridge | Extractive | Emits bounded chunks, Markdown-friendly tables, image placeholders, source metadata, and parser/conversion warnings. |
| <!-- capability:producer-corpus --> Producer corpus | Broad | Interoperability coverage includes real Word 16, reproducible Outlook 16 Save As and MAPI HTML-encapsulation files, a pinned external WordPad/RichEdit artifact, four LibreOffice regressions, external Google Docs, macOS TextEdit/RTFD, Epic EHI, CRM-workflow, and helpdesk-workflow artifacts, and a reproducibly generated GemBox.Document fixture. CRM and helpdesk files are workflow evidence, not vendor-native exports. |

Each producer entry identifies whether its bytes are redistributed, verified externally, generated reproducibly, or synthetic. This prevents a grammar sample from being mistaken for producer interoperability. The [producer manifest](../OfficeIMO.Rtf.Tests/Documents/RtfCorpus/corpus-manifest.json) records exact provenance; `Build/Test-RtfExternalProducerEvidence.ps1` reproduces bounded read, web-safe HTML, Markdown, and diagnostic-preserving Word conversion for external samples.

## P2: editing and scale

| Capability | Status | Contract and boundary |
| --- | --- | --- |
| <!-- capability:semantic-editing --> Semantic editing | Full | Clone; block insert/remove/move; paragraph/table/image insertion; cross-run text replacement; and cross-paragraph bookmark replacement are native `RtfDocument` operations. |
| <!-- capability:lossless-structural-editing --> Lossless structural editing | Broad | Root syntax fragments/nodes can be inserted, removed, or moved; pictures and destination-group content such as headers/footers can be replaced without normalizing unrelated syntax. |
| <!-- capability:document-merge --> Semantic document merge | Broad | `AppendDocument` clones and remaps fonts, colors, revision authors, styles, list definitions/overrides, blocks, tables, and notes. `PreserveSections = true` retains source sections, page defaults, and running stories; the default flattens section boundaries and reports discarded stories. Conflicting document metadata and unsupported resources produce diagnostics. |
| <!-- capability:fuzz-properties --> Seeded fuzz/property lane | Full | Valid exact round trip, malformed groups, extreme control parameters, Unicode fallback widths, binary lengths/limits, and semantic normalization run on all RTF test targets. |
| <!-- capability:benchmark-budgets --> Performance and allocation budgets | Full | BenchmarkDotNet covers scale comparison; isolated probes enforce elapsed, allocation, peak-working-set, output-size, and corpus-size ceilings for core plus every adapter. The enforced operations include untrusted parsing, web-safe HTML, and a real producer-scale GemBox fixture in addition to synthetic small, medium, and large documents. |
| <!-- capability:conversion-docs --> Living documentation | Full | This matrix, the capability manifest, package READMEs, safe/strict recipes, workflow examples, and benchmark commands are checked into the owning repository. |

## Adapter fidelity summary

| Feature | Word | HTML | Markdown | PDF export | Reader |
| --- | --- | --- | --- | --- | --- |
| Paragraphs and rich runs | Broad | Broad | Broad | Broad | Extractive |
| Links and bookmarks | Broad | Broad with URL policy | Broad | Broad | Extractive |
| Styles and lists | Broad | Broad | Broad | Partial visual mapping | Extractive |
| Tables and merged cells | Broad | Broad | Broad | Broad | Extractive |
| Nested tables | Broad | Broad | Flattened + diagnostic | Flattened + diagnostic | Flattened + diagnostic |
| Images | Broad | Resolver or trusted data URI | Export callback | PNG/JPEG/DIB; WMF/EMF callback | Visual placeholder |
| Notes | Broad | Broad | Markdown footnotes | Appended semantic notes | Extractive |
| Headers/footers | Broad | Broad | Omitted + diagnostic | Broad text mapping | Extractive |
| Section layout | Broad | Metadata plus CSS subset | Flattened + diagnostic | Page-starting geometry and equal-width columns; named limits | Extractive |
| Objects/shapes | Omitted + report where unsupported | Metadata/text subset | Omitted + diagnostic | Text/image subset + report | Text/placeholder subset |

## Safe and strict recipe

```csharp
RtfReadOptions limits = RtfReadOptions.CreateUntrustedProfile();
using FileStream input = File.OpenRead("upload.rtf");
RtfReadResult read = await RtfDocument.LoadResultAsync(input, limits, cancellationToken: cancellationToken);

var htmlOptions = RtfToHtmlOptions.CreateWebSafeProfile();
RtfToHtmlResult converted = read.Document.ToHtmlResult(htmlOptions);
string html = converted.Value;

var report = new RtfConversionReport();
report.AddReadDiagnostics(read.Diagnostics, "upload.rtf");
report.Merge(converted.RtfReport);
report.RequireNoLoss();
```

Use strict mode when the workflow must reject any degradation. For display or extraction workflows, inspect `Diagnostics` and decide which explicit flatten/omit/block actions are acceptable.

## Performance commands

```powershell
# Validate every benchmark case without collecting publication measurements.
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Rtf.Benchmarks\OfficeIMO.Rtf.Benchmarks.csproj -- --filter "*" --job Dry --noOverwrite

# Enforce the committed small/medium/large regression ceilings.
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Rtf.Benchmarks\OfficeIMO.Rtf.Benchmarks.csproj -- --verify-budgets --json .\artifacts\rtf-budget-report.json

# Recheck the external producer records without redistributing their source bytes.
pwsh .\Build\Test-RtfExternalProducerEvidence.ps1

# Collect focused BenchmarkDotNet measurements.
dotnet run -c Release --framework net8.0 --project .\OfficeIMO.Rtf.Benchmarks\OfficeIMO.Rtf.Benchmarks.csproj -- --filter "*RtfCoreBenchmarks*" --job Short --noOverwrite
```

Budget ceilings are regression safeguards, not competitor claims. Use the same runtime, corpus, and hardware when comparing historical results.
