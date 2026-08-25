# OfficeIMO.Reader extraction performance

This is a Windows baseline for the public Reader facade after removing repeated format detection, eliminating the discarded rich Markdown projection graph, scaling XML sibling paths linearly, and validating YAML limits during its single model-loading parse. It is regression evidence for equivalent OfficeIMO work, not a claim that a broad text extractor performs the same contract.

## Environment and provenance

- AMD Ryzen 9 9950X3D2, 32 logical / 16 physical cores
- Windows 11 25H2, build 26200.9168
- .NET SDK 10.0.111
- .NET 8.0.30 x64 runtime
- BenchmarkDotNet 0.15.8 `ShortRun`: one launch, three warmups, three measured iterations
- routing and first Markdown indexing source commit `9c3677821fa5a4f5de7f519daf0a2d7c897738df`
- intermediate Markdown bound-span projection source commit `5215409a568d68aa7a2efef290c96c5f64fca90d`
- final lightweight Markdown projection source commit `89b63d9fc522f10bb91830c38377452c491c5c96`, clean worktree for the final Markdown, ZIP, and isolated evidence runs
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

The initial full matrix found repeated format detection when a rich document fallback called the public chunk API again. The table-heavy Markdown adapter then called `FindFinalNodeForAssociatedObject` once per block; each call enumerated the complete final syntax tree from its root. Indexing those reference-equality associations once removed the O(n²) traversal, and reading the `SourceSpan` already bound to built-in Markdown objects removed the remaining lookup index.

Reader still paid for a complete source-backed syntax tree, rich table-cell navigation graph, and object-tree binding that it immediately discarded. The owning Markdown parser now provides an internal transient projection parse that preserves top-level block spans and simple table values while omitting those unused graphs. Configured document transforms retain the full source-backed path, and nonstandard custom blocks trigger a full-parse fallback in the thin Reader adapter.

| Public `ReadDocument` lane | Initial | Indexed syntax lookup | Bound-span projection | Lightweight projection current | Total change |
| --- | ---: | ---: | ---: | ---: | ---: |
| Markdown, 80 sections and 80 five-row tables | 32.822 ms / 47.68 MB | 15.825 ms / 13.32 MB | 4.839 ms / 12.60 MB | 1.169 ms / 2.04 MB | 96.4% less time / 95.7% less allocation |
| ZIP, 20 copies of the Markdown corpus | 862.923 ms / 961.96 MB | 441.272 ms / 271.99 MB | 163.287 ms / 257.47 MB | 32.402 ms / 46.26 MB | 96.2% less time / 95.2% less allocation |

The heading/table projection lane moved from 15.378 ms / 13.18 MB initially and 5.980 ms / 12.45 MB after bound-span navigation to 1.303 ms / 1.89 MB. The direct Markdown evidence case emits the same 24,986 normalized bytes before and after the lightweight projection change, with SHA-256 `6392082e220dea1b8797c5b5b255b1da1ee8c28836aa9e7816d845265e8290c9`. Its isolated allocation fell from 13,283,960 to 2,210,144 bytes and managed peak growth from 13,328,256 to 2,224,784 bytes, while retained result memory stayed near 0.35 MB.

The aggregate ZIP lane remains an optimization target: 46.26 MB is approximately 20 complete Markdown results and is not contender evidence against raw archive traversal. Its three-operation `ShortRun` timing has a wide confidence interval; allocation and contract validation are the more stable gates for that lane.

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
| DOCX | 7/7 | 913,392 | 11,640 | 921,384 | 40,960 | 19,012 / 220 |
| XLSX | 4/4 | 809,160 | 6,456 | 823,888 | 77,824 | 2,776 / 124 |
| PPTX | 7/7 | 2,310,400 | 10,280 | 2,337,704 | 28,672 | 25,366 / 191 |
| HTML | 7/7 | 1,590,896 | 13,496 | 629,784 | 65,536 | 515 / 314 |
| Markdown, 80 sections and 80 five-row tables | 6/6 | 2,210,144 | 358,480 | 2,224,784 | 200,704 | 25,947 / 24,986 |
| CSV | 3/3 | 221,704 | 4,952 | 181,824 | 65,536 | 60 / 96 |
| JSON, 1,200 records | 3/3 | 3,918,720 | 1,410,768 | 3,799,032 | 135,168 | 58,993 / 122,099 |
| XML, 1,200 records | 3/3 | 7,699,312 | 2,224,576 | 7,748,760 | 499,712 | 86,637 / 198,141 |
| YAML, 1,200 records | 3/3 | 8,594,728 | 1,394,320 | 8,614,656 | 372,736 | 61,389 / 122,099 |
| MSG | 6/6 | 485,184 | 12,072 | 508,424 | 65,536 | 10,752 / 434 |
| EPUB | 4/4 | 580,000 | 15,840 | 608,800 | 65,536 | 1,395 / 113 |
| ZIP | 6/6 | 537,504 | 6,248 | 563,800 | 81,920 | 377 / 168 |

The complete evidence also includes PDF and malformed-PDF policy checks because they are part of the existing format-neutral corpus; no PDF-owner optimization or competitive conclusion is derived from those cases here.

## Comparison boundary

No general .NET comparison is accepted for the complete Reader result contract yet. A fair lane must materialize equivalent normalized text, rich tables, links, assets, source locations, diagnostics, bounded malformed-input behavior, and nested-container results. Optional external direct-process runners remain available for tools that can meet a declared subset, but their process duration and peak working set are reported independently rather than turned into a misleading ratio.

The contender policy remains at most 2x for equivalent elapsed time and allocation. This run removes two internal O(n²) paths, brings the equivalent YAML/JSON structured lane inside that boundary, and establishes complete memory/size measurement. It does not declare the Reader facade finished without an equivalent external competitor lane.
