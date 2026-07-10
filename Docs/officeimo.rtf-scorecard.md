# OfficeIMO RTF scorecard

Last verified: 2026-07-10

This scorecard tracks evidence for the [RTF support matrix](officeimo.rtf-support-matrix.md) and the [end-to-end market audit](reviews/officeimo.rtf-end-to-end-market-gap-2026-07-10.md). Green tests prove current contracts; they do not by themselves prove interoperability with external RTF producers.

## Current verified baseline

```powershell
dotnet test OfficeIMO.Rtf.Tests\OfficeIMO.Rtf.Tests.csproj -c Release --logger "console;verbosity=minimal"
```

| Target | Passed | Failed | Skipped |
| --- | ---: | ---: | ---: |
| `net10.0` | 547 | 0 | 0 |
| `net8.0` | 547 | 0 | 0 |
| `net472` | 547 | 0 | 0 |

Source/test inventory:

| Evidence | Current value |
| --- | ---: |
| RTF implementation C# files across core and adapters | 337 |
| Focused RTF test C# files | 68 |
| Test methods before theory expansion | 490 |
| `.rtf` corpus fixtures | 9 |
| External producer families represented in corpus | 0 |
| Dedicated RTF benchmark projects | 0 |

## Product readiness

| Contract | Evidence | Status | Priority |
| --- | --- | --- | --- |
| Published core and adapter packages | Five RTF packages available from NuGet | Proven | Maintain |
| Multi-target build/test | 547 passing on three tested targets | Proven | Maintain |
| Lossless source round trip | Syntax tree and lossless editor tests | Proven for focused fixtures | Expand corpus |
| Rich semantic read/write | Focused model/writer tests across text, tables, fields, notes, sections, objects, shapes, metadata | Broad | Expand spec coverage |
| Safe untrusted RTF profile | Only `MaxDepth` exists in the core | Missing | P0 |
| Bounded binary/image/object payloads | No core limits | Missing | P0 |
| RTF-to-HTML safe URL policy | Output has no `HtmlUrlPolicy` | Missing | P0 |
| Unknown destination diagnostics | Unknown ignorable destinations can be skipped silently | Incomplete | P0 |
| Shared conversion report and strict mode | Fragmented reports; Word has none | Missing | P0 |
| Real producer interoperability corpus | Nine synthetic fixtures only | Missing | P1 |
| Word style and list structure parity | No Word style/list structure mapping | Missing | P1 |
| Word object/shape loss reporting | Blocks/inlines can be ignored with no report | Missing | P1 |
| Outlook HTML-encapsulated RTF | No semantic controls found | Missing | P1 |
| East Asian/DBCS text | No 932/936/949/950 semantic decoding | Missing | P1 |
| Nested tables | No nested-table control support found | Missing | P1 |
| PDF package | Published bidirectional semantic converter | Proven as semantic/extractive | Correct docs |
| PDF object/shape fidelity diagnostics | Plain-text fallback has no warning | Incomplete | P1 |
| Markdown notes/media workflows | Diagnostics exist; RTF notes omitted and media callback incomplete | Partial | P1 |
| Reader extraction | Block-aware chunks, diagnostics, and Reader input limits | Broad | Prove with corpus |
| Structural lossless edits | Targeted settings/metadata plus append | Partial | P2 |
| RTF-native semantic clone/merge/range edits | No public contract | Missing | P2 |
| Mail merge/find/replace/field/compare workflows | Engines exist in `OfficeIMO.Word`; no result-bearing RTF route | Partial | P2 after Word diagnostics |
| Benchmarks, memory budgets, and fuzzing | No dedicated lanes found | Missing | P2 |

## Adapter evidence

### Word

Proven focused contracts include rich run formatting, paragraph layout, tables and merged cells, images, notes, bookmarks, fields, comments, revisions, metadata, sections, page setup, headers/footers, and async I/O.

Open proof gaps:

- stylesheet and numbering structure preservation;
- shapes, objects, text boxes, and other unsupported elements;
- a conversion result that reports every omission or flattening;
- external Word-generated fixtures and reopen validation.

### HTML

Proven focused contracts include semantic text and formatting, tables, lists, images, fields, notes, revisions, metadata, document settings, sections, objects, shapes, and OfficeIMO round-trip attributes.

Open proof gaps:

- safe output URL schemes;
- web-safe versus round-trip profiles;
- bounded object/private metadata;
- caller-controlled image extraction/resource handling;
- browser validation of hostile and real producer files.

### Markdown

Proven focused contracts include headings, formatting, links, lists and restarts, tasks, tables, code blocks, definition lists, HTML fallbacks, Markdown-to-RTF footnotes, escaping, and diagnostics.

Open proof gaps:

- RTF notes to Markdown footnote definitions;
- headers/footers policy;
- complete binary image extraction/embedding callbacks;
- real producer fixtures.

### PDF

Proven focused contracts include semantic text, page setup, sections, lists, tables/merges, images, bookmarks, links, metadata, headers/footers, notes, file/stream/async APIs, diagnostics, and extractive PDF import.

Open proof gaps:

- object and shape fidelity reporting;
- page-native notes;
- reusable EMF/WMF/DIB conversion;
- font substitution/embedding diagnostics;
- target visual corpus and page-image workflows.

### Reader

Proven focused contracts include registration, chunking, tables, image placeholders, metadata/provenance, parser warnings, and Reader-level input byte limits.

Open proof gaps:

- bounded core RTF parsing shared by every adapter;
- real EHR/CRM/helpdesk and Outlook inputs;
- explicit revision/hidden/object ingestion policies.

## Next acceptance gates

### Gate 1 - Safe defaults

- [ ] Web-safe RTF-to-HTML profile rejects executable/unsafe links.
- [ ] Core untrusted profile bounds input, tokens, semantic objects, and binary payloads.
- [ ] Every limit emits a stable diagnostic and honors cancellation.
- [ ] Unknown ignorable destinations are reported when semantic content is discarded.

### Gate 2 - Diagnostic truth

- [ ] Shared RTF conversion report exists.
- [ ] Word conversion returns result plus diagnostics.
- [ ] Every adapter reports block/inline fallback or omission.
- [ ] Strict no-loss mode is covered by tests.

### Gate 3 - Interoperability proof

- [ ] Producer manifest and redistribution provenance are checked in.
- [ ] Word, LibreOffice, Outlook, Google Docs, and macOS fixtures exist.
- [ ] Lossless, semantic, normalized, target-reopen, and visual expectations are recorded.
- [ ] Matrix `Full` claims are generated from or verified against this evidence.

### Gate 4 - Workflow parity

- [ ] Word styles and numbering survive the RTF bridge or produce diagnostics.
- [ ] Objects/shapes are mapped or explicitly diagnosed.
- [ ] Thin RTF mail-merge, find/replace, field-update, compare, and merge workflows call `OfficeIMO.Word`.
- [ ] Output includes combined parse, conversion, and workflow diagnostics.

### Gate 5 - Scale and product polish

- [ ] Benchmarks cover parse, lossless write, semantic write, adapters, and Reader.
- [ ] Allocation and peak-memory budgets exist for representative sizes.
- [ ] Fuzz/property tests cover malformed syntax, Unicode, binary lengths, and lossless invariants.
- [ ] NuGet documentation shows safe ingestion, strict conversion, and the current conversion graph.
