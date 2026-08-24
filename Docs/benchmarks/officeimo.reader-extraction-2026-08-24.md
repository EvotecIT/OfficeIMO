# OfficeIMO.Reader extraction performance

This is a Windows baseline for the public Reader facade after removing repeated format detection, indexing Markdown syntax-to-model navigation, scaling XML sibling paths linearly, and validating YAML limits during its single model-loading parse. It is regression evidence for equivalent OfficeIMO work, not a claim that a broad text extractor performs the same contract.

## Environment and provenance

- AMD Ryzen 9 9950X3D2, 32 logical / 16 physical cores
- Windows 11 25H2, build 26200.9168
- .NET SDK 10.0.111
- .NET 8.0.30 x64 runtime
- BenchmarkDotNet 0.15.8 `ShortRun`: one launch, three warmups, three measured iterations
- routing and first Markdown indexing source commit `9c3677821fa5a4f5de7f519daf0a2d7c897738df`
- final Markdown bound-span projection source commit `5215409a568d68aa7a2efef290c96c5f64fca90d`, clean worktree for the final Markdown and ZIP benchmark run
- final structured-format source commit `18fc5f2b045a4b6272d5cff1743b3edb1a4d407f`, clean worktree for final isolated extraction evidence
- deterministic benchmark corpus; source creation is outside measured operations; source hashing is disabled

Commands:

```powershell
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Reader.Benchmarks -- --filter '*ReaderDocumentBenchmarks*'
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Reader.Benchmarks -- --filter '*ReaderMarkdownPipelineBenchmarks*'
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Reader.Benchmarks -- --filter '*ReaderXmlSiblingBenchmarks*'
dotnet run -c Release -f net8.0 --project .\OfficeIMO.Reader.Benchmarks -- evidence --output .benchmark-artifacts\reader-evidence
```

## Public extraction improvement

The initial full matrix found repeated format detection when a rich document fallback called the public chunk API again. The table-heavy Markdown adapter then called `FindFinalNodeForAssociatedObject` once per block; each call enumerated the complete final syntax tree from its root. Indexing those reference-equality associations once removed the O(n²) traversal. The final adapter now reads the `SourceSpan` already bound to built-in Markdown objects and retains the syntax-tree lookup only for custom blocks outside that object model, avoiding even the one complete lookup-index construction.

| Public `ReadDocument` lane | Initial | Indexed syntax lookup | Bound-span current | Total change |
| --- | ---: | ---: | ---: | ---: |
| Markdown, 80 sections and 80 five-row tables | 32.822 ms / 47.68 MB | 15.825 ms / 13.32 MB | 4.839 ms / 12.60 MB | 85.3% less time / 73.6% less allocation |
| ZIP, 20 copies of the Markdown corpus | 862.923 ms / 961.96 MB | 441.272 ms / 271.99 MB | 163.287 ms / 257.47 MB | 81.1% less time / 73.2% less allocation |

The heading/table projection lane moved from 15.378 ms / 13.18 MB to 5.980 ms / 12.45 MB. Parsing with the required syntax tree and tables still accounts for 11.01 MB, so the remaining absolute cost is primarily the rich public parse model rather than navigation. The aggregate ZIP lane remains an optimization target: its 257.47 MB is approximately 20 complete Markdown results and is not contender evidence against raw archive traversal. Its three-operation `ShortRun` timing also has a wide confidence interval; allocation and contract validation are the more stable gates for that lane.

Short-run timings for other routed formats are intentionally not promoted as release claims because the samples are narrow and some confidence intervals are wide. Their allocation remained stable while EPUB, HTML, Word, XML, YAML, Excel, and ZIP no longer repeat the same detection route.

## Structured-format improvements

XML previously calculated every element ordinal with `ElementsBeforeSelf(...).Count()`. Repeated same-name siblings therefore became quadratic even though the output path contract only needs a per-parent occurrence count. The dedicated scaling lane now assigns ordinals in one pass and validates alternating names plus a fifth-name overflow path.

| Repeated XML siblings | Before mean / allocation | After mean / allocation | Change |
| ---: | ---: | ---: | ---: |
| 128 | 1.397 ms / 830.63 KB | 1.020 ms / 806.56 KB | 27.0% less time / 2.9% less allocation |
| 1,200 | 6.290 ms / 6,445.51 KB | 3.533 ms / 6,220.47 KB | 43.8% less time / 3.5% less allocation |
| 10,000 | 249.638 ms / 53,158.12 KB | 24.252 ms / 51,282.94 KB | 90.3% less time / 3.5% less allocation |

YAML previously parsed every input twice: a streaming security preflight followed by YamlDotNet representation-model loading. An `IParser` wrapper now enforces the same event, node, depth, and scalar limits while the model consumes its one event stream. Already-normalized keys and scalars also avoid redundant `StringBuilder` copies. All 17 YAML contract and limit tests continue to pass.

The final paired 1,200-record lanes produce the same normalized 122,099-byte output shape. YAML is now within the contender boundary against JSON for this equivalent Reader contract:

| Format | Mean | Allocated | Ratio to JSON |
| --- | ---: | ---: | ---: |
| JSON | 3.548 ms | 5.19 MB | 1.00x / 1.00x |
| YAML | 6.569 ms | 9.74 MB | 1.85x time / 1.88x allocation |

Before these changes, the YAML lane measured 12.341 ms / 12.95 MB. The final result is 46.8% faster and allocates 24.8% less while retaining fail-closed parsing limits.

## Isolated memory and output evidence

The `evidence` command validates each normalized extraction twice, checks format-neutral Markdown plus OfficeIMO-native rich probes, and then starts one child process per case for retained and peak-memory sampling. Output bytes are normalized extracted Markdown bytes; they are not source-format round-trip sizes.

| Case | Probes | Allocated | Retained | Managed peak growth | Process working-set growth | Input / output bytes |
| --- | ---: | ---: | ---: | ---: | ---: | ---: |
| DOCX | 7/7 | 907,120 | 10,632 | 913,400 | 69,632 | 19,080 / 220 |
| XLSX | 4/4 | 789,984 | 5,448 | 802,784 | 81,920 | 2,774 / 124 |
| PPTX | 7/7 | 2,291,552 | 9,272 | 2,320,592 | 53,248 | 25,342 / 191 |
| HTML | 7/7 | 1,587,184 | 12,488 | 628,712 | 69,632 | 515 / 314 |
| CSV | 3/3 | 219,864 | 3,944 | 181,816 | 81,920 | 60 / 96 |
| JSON, 1,200 records | 3/3 | 3,924,912 | 1,401,536 | 3,800,240 | 188,416 | 58,993 / 122,099 |
| XML, 1,200 records | 3/3 | 7,702,616 | 2,223,856 | 7,750,280 | 757,760 | 86,637 / 198,141 |
| YAML, 1,200 records | 3/3 | 8,600,408 | 1,401,536 | 8,623,776 | 335,872 | 61,389 / 122,099 |
| MSG | 6/6 | 483,520 | 11,064 | 514,112 | 77,824 | 10,752 / 434 |
| EPUB | 4/4 | 577,320 | 14,736 | 607,208 | 77,824 | 1,395 / 113 |
| ZIP | 6/6 | 552,848 | 5,240 | 584,496 | 65,536 | 377 / 168 |

The complete evidence also includes PDF and malformed-PDF policy checks because they are part of the existing format-neutral corpus; no PDF-owner optimization or competitive conclusion is derived from those cases here.

## Comparison boundary

No general .NET comparison is accepted for the complete Reader result contract yet. A fair lane must materialize equivalent normalized text, rich tables, links, assets, source locations, diagnostics, bounded malformed-input behavior, and nested-container results. Optional external direct-process runners remain available for tools that can meet a declared subset, but their process duration and peak working set are reported independently rather than turned into a misleading ratio.

The contender policy remains at most 2x for equivalent elapsed time and allocation. This run removes two internal O(n²) paths, brings the equivalent YAML/JSON structured lane inside that boundary, and establishes complete memory/size measurement. It does not declare the Reader facade finished without an equivalent external competitor lane.
