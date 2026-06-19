# OfficeIMO RTF Scorecard

This scorecard tracks evidence for the RTF market roadmap. It is intentionally evidence-first: green cells require current tests or fixtures, not assumptions.

## Current Verified Baseline

Command last run from `C:\Support\GitHub\OfficeIMO-rtf-market-assessment-20260618`:

```powershell
dotnet test OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --filter "FullyQualifiedName~Rtf"
```

Result:

| Target | Passed | Failed | Skipped |
| --- | ---: | ---: | ---: |
| `net10.0` | 551 | 0 | 0 |
| `net8.0` | 551 | 0 | 0 |
| `net472` | 551 | 0 | 0 |

PDF work is intentionally out of scope for the current implementation phases. The current non-PDF RTF validation command is:

```powershell
dotnet test OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --no-restore --filter "FullyQualifiedName~Rtf&FullyQualifiedName!~Pdf"
```

Result:

| Target | Passed | Failed | Skipped |
| --- | ---: | ---: | ---: |
| `net10.0` | 523 | 0 | 0 |
| `net8.0` | 523 | 0 | 0 |
| `net472` | 523 | 0 | 0 |

## Phase 1 Contract Evidence

| Deliverable | Evidence | Status |
| --- | --- | --- |
| Support matrix | `Docs/officeimo.rtf-support-matrix.md` | Added |
| Conversion class vocabulary | `Lossless`, `Semantic`, `Visual`, `Extractive`, `Diagnostic` in support matrix | Added |
| Golden corpus structure | `OfficeIMO.Tests/Documents/RtfCorpus` | Added |
| Golden corpus parse guard | `RtfGoldenCorpusTests.GoldenCorpusFixturesParseWithoutErrors` | Passing |
| PDF deferral boundary | PDF rows marked `Deferred`; no PDF implementation edits | Added |

## Phase 2 RTF <=> Markdown Evidence

| Contract | Evidence | Status |
| --- | --- | --- |
| Direct RTF to Markdown text | `RtfMarkdownConverterTests.RtfDocumentToMarkdownPreservesCoreInlineBlocksListsAndTables` | Passing |
| Direct Markdown to RTF document | `RtfMarkdownConverterTests.MarkdownToRtfDocumentPreservesHeadingsInlinesListsAndTables` | Passing |
| Rich inline projection | Bold, italic, strike, underline, links, breaks | Passing |
| Table projection | Header/body tables and cell text | Passing |
| Diagnostics for unsupported rich constructs | Image/raw HTML/object diagnostics | Passing |

## Phase 3 Word/HTML Evidence

| Contract | Evidence | Status |
| --- | --- | --- |
| Word bridge leading note references | `WordRtfConverterTests.Word_To_Rtf_Bridge_Carries_Leading_Footnote_Reference` | Passing |
| RTF to HTML skipped-image diagnostics | `RtfHtmlConverterTests.RtfDocument_ToHtml_Reports_Diagnostic_When_Image_Embedding_Is_Disabled` | Passing |
| RTF to HTML unsupported image diagnostics | `RtfHtmlConverterTests.RtfDocument_ToHtml_Reports_Diagnostic_For_Unsupported_Image_Format` | Passing |
| RTF to HTML options diagnostics contract | `RtfHtmlOptionsTests.RtfToHtmlOptions_Clone_Copies_Configuration` | Passing |
| PDF implementation | Deferred to avoid cross-agent collision | Deferred until final phase |

Focused Phase 3 validation:

```powershell
dotnet test OfficeIMO.Tests\OfficeIMO.Tests.csproj -c Release --no-restore --filter "FullyQualifiedName~WordRtfConverterTests|FullyQualifiedName~RtfHtmlOptionsTests|FullyQualifiedName~RtfHtmlConverterTests"
```

Result:

| Target | Passed | Failed | Skipped |
| --- | ---: | ---: | ---: |
| `net10.0` | 82 | 0 | 0 |
| `net8.0` | 82 | 0 | 0 |
| `net472` | 82 | 0 | 0 |
