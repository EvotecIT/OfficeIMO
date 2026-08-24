# OfficeIMO.Reader extraction performance

This is a Windows baseline for the public Reader facade after removing repeated format detection and indexing Markdown syntax-to-model navigation. It is regression evidence for equivalent OfficeIMO work, not a claim that a broad text extractor performs the same contract.

## Environment and provenance

- AMD Ryzen 9 9950X3D2, 32 logical / 16 physical cores
- Windows 11 25H2, build 26200.9168
- .NET SDK 10.0.111
- .NET 8.0.30 x64 runtime
- BenchmarkDotNet 0.15.8 `ShortRun`: one launch, three warmups, three measured iterations
- optimized source commit `9c3677821fa5a4f5de7f519daf0a2d7c897738df`, clean worktree for isolated extraction evidence
- deterministic benchmark corpus; source creation is outside measured operations; source hashing is disabled

Commands:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Reader.Benchmarks -- --filter '*ReaderDocumentBenchmarks*'
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Reader.Benchmarks -- --filter '*ReaderMarkdownPipelineBenchmarks*'
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Reader.Benchmarks -- evidence --output .benchmark-artifacts\reader-evidence
```

## Public extraction improvement

The initial full matrix found repeated format detection when a rich document fallback called the public chunk API again. The table-heavy Markdown adapter then called `FindFinalNodeForAssociatedObject` once per block; each call enumerated the complete final syntax tree from its root. Indexing those reference-equality associations once removes the O(n²) traversal while preserving the first matching syntax node.

| Public `ReadDocument` lane | Before mean / allocation | After mean / allocation | Change |
| --- | ---: | ---: | ---: |
| Markdown, 80 sections and 80 five-row tables | 32.822 ms / 47.68 MB | 15.825 ms / 13.32 MB | 51.8% less time / 72.1% less allocation |
| ZIP, 20 copies of the Markdown corpus | 862.923 ms / 961.96 MB | 441.272 ms / 271.99 MB | 48.9% less time / 71.7% less allocation |

The isolated Markdown pipeline now allocates 13.09-13.18 MB instead of 47.68 MB. Parsing with the required syntax tree and tables accounts for 11.01 MB, so the remaining chunk projection overhead is roughly 2.1 MB rather than a second parser-sized traversal. The aggregate ZIP lane remains an optimization target; its cost is approximately 20 complete Markdown results and is not contender evidence against raw archive traversal.

Short-run timings for other routed formats are intentionally not promoted as release claims because the samples are narrow and some confidence intervals are wide. Their allocation remained stable while EPUB, HTML, Word, XML, YAML, Excel, and ZIP no longer repeat the same detection route.

## Isolated memory and output evidence

The `evidence` command validates each normalized extraction twice, checks format-neutral Markdown plus OfficeIMO-native rich probes, and then starts one child process per case for retained and peak-memory sampling. Output bytes are normalized extracted Markdown bytes; they are not source-format round-trip sizes.

| Case | Probes | Allocated | Retained | Managed peak growth | Process working-set growth | Input / output bytes |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| DOCX | 7/7 | 907,120 | 10,632 | 913,400 | 69,632 | 19,080 / 220 |
| XLSX | 4/4 | 789,984 | 5,448 | 802,784 | 81,920 | 2,774 / 124 |
| PPTX | 7/7 | 2,291,552 | 9,272 | 2,320,592 | 53,248 | 25,342 / 191 |
| HTML | 7/7 | 1,587,184 | 12,488 | 628,712 | 69,632 | 515 / 314 |
| CSV | 3/3 | 219,864 | 3,944 | 181,816 | 81,920 | 60 / 96 |
| MSG | 6/6 | 483,520 | 11,064 | 514,112 | 77,824 | 10,752 / 434 |
| EPUB | 4/4 | 577,320 | 14,736 | 607,208 | 77,824 | 1,395 / 113 |
| ZIP | 6/6 | 552,848 | 5,240 | 584,496 | 65,536 | 377 / 168 |

The complete evidence also includes PDF and malformed-PDF policy checks because they are part of the existing format-neutral corpus; no PDF-owner optimization or competitive conclusion is derived from those cases here.

## Comparison boundary

No general .NET comparison is accepted for the complete Reader result contract yet. A fair lane must materialize equivalent normalized text, rich tables, links, assets, source locations, diagnostics, bounded malformed-input behavior, and nested-container results. Optional external direct-process runners remain available for tools that can meet a declared subset, but their process duration and peak working set are reported independently rather than turned into a misleading ratio.

The contender policy remains at most 2x for equivalent elapsed time and allocation. This run removes an internal O(n²) incident and establishes complete memory/size measurement; it does not declare the Reader facade finished without an equivalent competitor lane.
