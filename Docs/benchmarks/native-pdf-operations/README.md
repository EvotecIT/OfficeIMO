# Native PDF operations: 8 September 2026

[The measured data](2026-09-08-windows-net8.json) compares product commits `d3aa130ac52beb39cf382cef7987cc5b29d5bd2e` and `9e4001e83179d39636edb3b976cd6befd381e40a` on .NET 8.0.30. BenchmarkDotNet 0.15.8 ran its full default job in separate baseline/current processes on two CPU-affinity domains of one Windows workstation. The JSON retains iteration values, statistical summaries, managed allocation, host information, and hashes of the original reports.

## Equivalent work

Each independent producer creates a 100-page PDF with four table rows and one narrative paragraph per page. The benchmark includes OfficeIMO source parsing and, for page operations, serialization of every output. Setup reopens inputs and outputs with PdfPig and requires every expected page and narrative/table value. Selection emits 25 pages in descending source order. Merge combines 25 four-page documents. Both split methods emit the same 100 single-page documents.

## Observations

Timing ranges below span the two affinity runs. Allocation values are managed MiB per complete operation; they are not peak heap or process memory.

| Operation and input | Baseline mean | Current mean | Baseline allocation | Current allocation |
| --- | ---: | ---: | ---: | ---: |
| Explicit split selections, iText | 518–742 ms | 34–35 ms | 592.57 MiB | 68.05 MiB |
| Explicit split selections, MigraDoc | 995–1,117 ms | 48–51 ms | 1,032 MiB | 85.79 MiB |
| Page selection, iText | 11.6–15.0 ms | 9.2–10.9 ms | 12.20 MiB | 8.77 MiB |
| Page selection, MigraDoc | 21.1–23.1 ms | 9.8–10.5 ms | 18.57 MiB | 14.08 MiB |
| Whole-document split, iText | 37.3–37.5 ms | 34.0–35.2 ms | 70.56 MiB | 68.03 MiB |
| Whole-document split, MigraDoc | 53.7–59.5 ms | 55.9–59.0 ms | 89.83 MiB | 85.77 MiB |
| Plain-text reading, iText | 40.6–45.6 ms | 39.2–41.3 ms | 63.43 MiB | 61.36 MiB |
| Plain-text reading, MigraDoc | 51.4–57.5 ms | 45.8–60.0 ms | 72.03 MiB | 69.00 MiB |
| Structured reading, iText | 181.6–187.4 ms | 177.7–179.4 ms | 165.94 MiB | 161.11 MiB |
| Structured reading, MigraDoc | 215.7–226.6 ms | 203.3–229.0 ms | 215.39 MiB | 208.31 MiB |
| Merge, iText | 31.0–31.2 ms | 31.1–31.3 ms | 37.93 MiB | 38.17 MiB |
| Merge, MigraDoc | 34.0–35.8 ms | 33.0–38.5 ms | 52.84 MiB | 52.95 MiB |

Reusing one extraction session removes repeated whole-source parsing from compound split selections. That gain is large in both runs. Page selection also consistently allocates less and completes sooner. The reader's borrowed operand buffer reduces allocation by roughly 3–4%; timings for reading and ordinary splitting are less conclusive. Merge already parsed each source once, and now takes a fresh metadata snapshot to prevent edits to a returned read model from changing the source document's page-operation metadata.

Several short builds overlapped the first run. The second used a different affinity domain without concurrent full builds, but neither run represents a dedicated laboratory host. Small changes, overlapping distributions, and differences that change sign across runs do not establish a general speedup or regression. In particular, this evidence does not claim faster merge or uniformly faster structured reading.

## Reproduce

Run the committed benchmark through the shared evidence runner:

```powershell
pwsh Build/Run-LibraryComparisonBenchmarks.ps1 -Workload pdfnative -RunMode full -Framework net8.0 -AffinityMask 65535
```

The recorded comparison used the same committed benchmark class and validation helpers linked into a small BenchmarkDotNet host. Its baseline job referenced release-built Core/Pdf assemblies from the baseline commit; its current job used the current Pdf project reference. Both jobs used `Job.Default`, `MemoryDiagnoser`, and identical producer/scenario parameters. Run the same workload at each selected revision when comparing later changes; do not compare these dated results with different page counts, excluded serialization, or unvalidated output.
