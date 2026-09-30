# Excel comparison snapshot, 28 September 2026

This focused comparison refreshes the four Excel operations shown in the package README and benchmark website. Each 25,000-row lane validates its output before timing. It is a local engineering comparison, not evidence of a source-level speedup or a portable ranking. The broader scenario matrix has not been refreshed by this run.

The Windows 11 host was an AMD Ryzen 9 9950X3D2 running .NET 8.0.31, with `Normal` process priority. Runs alternated between the complete 16-logical-processor domains `0xFFFF` (A1), `0xFFFF0000` (B1), `0xFFFF0000` (B2), and `0xFFFF` (A2). Each run used 20 warmups and nine measured iterations; the raw JSON retains every sample and allocation observation. A2 supplies the compact README and website snapshot. The runner assembly SHA-256 was `CC6BB4BEA0E52CB3BDB5A64E73201402B6529A9A8ED12EBCB956006F04D9B5B5`; the measured Excel assembly SHA-256 was `11B203E00339B33E77BB1912252ABD771726D52BA5116D6E3DDE1A2583BF70A2`. The source base was `185cdb3e2b2563f69ae1ba993c3225f49d02a3ca`, plus the benchmark-runner metadata change in this commit.

| OfficeIMO operation | A1 median | B1 median | B2 median | A2 median | A2 allocation |
| --- | ---: | ---: | ---: | ---: | ---: |
| Feature-rich report create/save | 53.74 ms | 42.07 ms | 39.56 ms | 52.11 ms | 16,582 KiB |
| Typed object read | 35.25 ms | 26.83 ms | 24.47 ms | 32.79 ms | 2,034 KiB |
| Styled DataReader table export | 29.04 ms | 21.45 ms | 21.34 ms | 28.48 ms | 6,122 KiB |
| Compact DataReader export | 27.14 ms | 22.94 ms | 21.41 ms | 26.65 ms | 6,061 KiB |

The selected A2 snapshot shows OfficeIMO faster than the measured peers for the report, typed read, and styled table lanes. In compact export, the runner's **means** treat OfficeIMO and SpreadCheetah as a tie within its 5% threshold (26.09 versus 26.76 ms), while the published **medians** are 26.65 versus 28.74 ms. The source [A2 summary](excel-readme-2026-09-28/a2/officeimo.excel.comparison-summary.md) gives all peer results, means, dispersion, and allocations. The two processor domains differ materially in elapsed time, while OfficeIMO's read and compact-write allocation is stable. A2's styled-export mean is 36.01 ms versus a 28.48 ms median because its raw samples include a slow outlier; neither that sample nor the slower A-domain runs were discarded.

Reproduce the comparison on a machine whose topology supports the selected mask:

```powershell
dotnet run -c Release -f net8.0 --project ./OfficeIMO.Excel.Benchmarks -- comparison-suite ./Ignore/Benchmarks/excel-readme-local --row-set 25000 --scenario realworld-report-all-in-one,write-datareader-table,read-objects-stream,write-datareader-compact-package --skip-package-profile --skip-dense-helloworld --skip-legacy-epplus --warmup 20 --iterations 9 --affinity 0xFFFF --priority Normal
```

Repeat with `0xFFFF0000` only if that mask selects another complete cache domain. The four [raw A1](excel-readme-2026-09-28/a1/officeimo.excel.comparison-speed-25000.json), [B1](excel-readme-2026-09-28/b1/officeimo.excel.comparison-speed-25000.json), [B2](excel-readme-2026-09-28/b2/officeimo.excel.comparison-speed-25000.json), and [A2](excel-readme-2026-09-28/a2/officeimo.excel.comparison-speed-25000.json) artifacts preserve sample-level evidence. A2 and B2 also retain complete suite manifests and generated summaries. The A1/B1 raw files were copied from scratch output after the run; their companion summaries were intentionally not published because they embedded the scratch path.

The selected compact artifact is [hash-pinned](readme-current/officeimo.excel.comparison.json) to the A2 summary. The generated benchmark page displays its date, operating system, affinity, priority, runtime, workload, and per-library result. The comparison is a bounded snapshot: full-matrix refresh, independent peak-memory budgets, additional sizes, and native Linux/macOS evidence remain open in the [roadmap](../ROADMAP.md#spreadsheet-and-csv-delivery-order).
