# XLSX buffering, CSV escaping, and reader allocation measurements — 2026-09-08

This run measures another round of changes after the September 7 text-serialization baseline. Complete XLSX exports batch UTF-8 encoding and stream writes. CSV exports batch dense quote escaping. XLSX reads avoid allocating unused caches and reduce the cost of parsing small package metadata.

Public APIs and the decoded file contract remain unchanged. The XLSX encoder can change ZIP compression boundaries, so compressed package bytes need not be identical; every uncompressed package part was compared across eight export profiles.

## Measurement boundaries

The before version is the September 7 candidate at commit `3f2b32a60114514e92af24ba682aae2d4ca3e097`; its product assemblies match the saved phase-one binaries. This report measures additional changes against that baseline. The [earlier report](officeimo.excel-csv-text-2026-09-07.md) covers the first round against the original source.

The measured workstation has an AMD Ryzen 9 9950X3D2, 16 cores and 32 logical processors, with two 96 MiB L3 domains. Measurements use .NET 10.0.11, Windows, Normal priority, and one cache domain at a time. Other work continued on the machine. No outliers were removed.

The final writer comparison uses six fresh PowerShell processes: three repetitions on each domain, with 12 warmups, 30 measured samples, and eight complete exports per sample. PowerForge rotates before/after order. Identical benchmark and dependency assemblies load both product versions in separate assembly contexts. Each setup validates complete output; file cases also compare both writers' UTF-8 bytes and decoded fields. File writes include opening, closing, and flushing to the operating system, without forcing durable storage.

The identical-binary calibration repeats the same baseline on both sides in four fresh processes, with 16 exports per sample. It shows that this workstation and harness can report apparent differences above 5% even when the product binaries are identical. Calibration values are retained separately, without subtracting them from candidate results.

## Complete export results

The table reports pooled medians across all three repetitions on each domain. Negative changes mean less elapsed time. [All final writer samples](excel-csv-buffering-2026-09-08/writers.json) include means and every process result as well as medians.

| Complete workload | L3 domain 0: before → after | L3 domain 1: before → after | Median change |
| --- | ---: | ---: | ---: |
| XLSX long plain, compact | 1.719 → 1.598 ms | 1.533 → 1.377 ms | -7.0% / -10.2% |
| XLSX long escaped, compact | 2.985 → 2.627 ms | 2.371 → 2.230 ms | -12.0% / -6.0% |
| XLSX long markup, compact | 11.949 → 8.279 ms | 9.587 → 7.114 ms | -30.7% / -25.8% |
| XLSX long plain, public defaults | 3.224 → 2.965 ms | 2.905 → 2.818 ms | -8.0% / -3.0% |
| XLSX long escaped, public defaults | 4.351 → 4.355 ms | 3.886 → 3.749 ms | +0.1% / -3.5% |
| XLSX 25,000-row control | 19.589 → 18.322 ms | 16.293 → 15.629 ms | -6.5% / -4.1% |
| CSV long quote runs, in memory | 4.874 → 1.686 ms | 3.409 → 1.422 ms | -65.4% / -58.3% |
| CSV long quote runs, UTF-8 file | 8.206 → 5.215 ms | 7.209 → 4.976 ms | -36.4% / -31.0% |
| CSV short quote runs, AsNeeded | 0.317 → 0.278 ms | 0.288 → 0.265 ms | -12.3% / -8.1% |
| CSV short quote runs, Always | 0.367 → 0.295 ms | 0.267 → 0.225 ms | -19.5% / -15.7% |
| CSV long JSON, in memory | 3.710 → 3.168 ms | 3.064 → 2.929 ms | -14.6% / -4.4% |
| CSV long notes, in memory | 1.147 → 1.118 ms | 0.987 → 0.969 ms | -2.6% / -1.9% |
| CSV long notes, UTF-8 file | 5.425 → 5.235 ms | 4.774 → 4.873 ms | -3.5% / +2.1% |
| CSV 25,000-row control | 6.405 → 6.204 ms | 5.458 → 5.523 ms | -3.1% / +1.2% |

The long-plain XLSX target was another 25–30% reduction. The observed 7–10% reduction falls short. Managed sampling identifies the writer flush and output-buffer paths, but its inlining attribution does not justify a precise CPU split between memory copies, encoding, and native compression. Long markup benefits more from accumulating small XML fragments before encoding. Public defaults retain shared strings and explicit cell references, and their results are reported separately from compact streaming.

The initial [baseline qualification](excel-csv-buffering-2026-09-08/qualification-summary.json) repeated both 25,000-row controls in six fresh processes against the original source. Every process median was within 4.5%, so the earlier isolated XLSX control slowdown did not repeat. Dense-text timing varied, motivating the later quote-escaping work.

Consistency is not uniform across short CSV cases:

- Short JSON under `Always` has a median change of about +0.5% on both domains. Under `AsNeeded`, domain 0 is +7.9% by median and +7.3% by mean, while domain 1 is -4.3% by median and effectively unchanged by mean. The domain-0 process mean changes are +10.5%, +11.8%, and -0.3%; this possible regression remains visible in the data.
- Short ASCII files are +5.4% / +2.3% by median and +4.0% / +0.4% by mean. The identical-binary calibration itself reports +8.8% / +11.7% means for that case. Its ordinary text does not enter the changed quote-escaping helper.
- Typed file writes under `AsNeeded` are +1.3% / +9.6% by median, but -2.2% / +1.7% by mean. The process results and identical-binary calibration vary in direction. These measurements do not establish a reliable gain or regression.
- Short Unicode and custom-delimiter files are within 2.5% by median. Long-note file means are -1.9% / +1.2%, preserving the earlier improvement without an additional consistent gain.

The [identical-binary calibration](excel-csv-buffering-2026-09-08/identical-binary-calibration.json) exposes measurement uncertainty. It does not erase an unfavorable candidate result or prove that small regressions are impossible.
## Reader allocation

The XLSX fixture contains 65,535 rows and 14 columns. The measured scan opens the package, reads typed values, and validates its observation. Small package parts remain subject to declared-length checks, byte limits, cancellation, and URI normalization. Pooled bytes are cleared when returned. Larger XML parts retain the BCL streaming reader.

The final native `MemoryDiagnoser` comparison reports **183,832 → 116,720 bytes per XLSX scan**, a **36.5% reduction** on both domains. The 50% allocation target is not met. Remaining allocations include BCL XML-reader state, parsed metadata objects, shared-string values, and ZIP entry indexing. Allocation traces also contain initial cold pool rentals; those are not counted as steady per-scan allocation in this result.

The separate [rotated reader comparison](excel-csv-buffering-2026-09-08/reader-timing.json) has pooled mean changes of +0.9% / -1.4%, with median changes of -1.5% / +0.8%. The native sequential comparison is +3.2% / +1.4% by mean. These observations do not establish a repeatable elapsed regression, but do not exclude a small one. Reader source was frozen for the rotated run; its timing preceded the final writer-only refinements, while native allocation uses the final product assemblies.

The native short-JSON confirmation changes mean time by -3.0% / +2.9%. It does not reproduce the larger one-domain slowdown in the rotated matrix. Both observations remain in the evidence; short-field consistency is still an open optimization target.
CSV uses the existing string-returning scan. Independent fixture analysis counts 917,490 fields and 7,253,195 UTF-16 characters. Assuming ordinary x64 .NET string layout, allocating a fresh string for every field would cost 37,441,976 bytes. The parser already reuses 65,593 one-character ASCII values, reducing the estimated fresh field-string cost to 35,867,744 bytes. Native allocation is **35,873,104 bytes** for both CSV versions, leaving about **5,360 bytes** beyond the estimated fresh field strings. The estimate depends on this fixture and API; it is not a general minimum for every CSV reader. There are no CSV reader production changes. Native sequential elapsed time varies by +14.7% / -2.2% despite identical read code and allocation, another reason to retain the busy-host uncertainty.

## Native library comparisons

Native BenchmarkDotNet runs cover 48 in-memory CSV cases, 48 UTF-8 file cases, 24 compact XLSX cases, and 12 public-default XLSX cases, plus 12 before/after reader and short-JSON cases. Each has 16 measured iterations with all outliers retained. The comparison packages are CsvHelper 33.1.0 and SpreadCheetah 1.28.0; independent XLSX validation uses ExcelDataReader 3.9.0. Managed allocations are measured; this run does not establish peak native-memory or Linux/macOS budgets. The selected comparison below reports OfficeIMO mean time divided by the other library's mean; a value below 1 means OfficeIMO used less time. Allocation values average the two domains and include the result retained by each workload.

| Workload | Comparison library | OfficeIMO / comparison mean time, L3 0 / L3 1 | OfficeIMO allocation / comparison allocation |
| --- | --- | ---: | ---: |
| CSV short JSON, text | CsvHelper | 1.65× / 1.58× | 346.2 / 618.2 KiB |
| CSV long notes, text | CsvHelper | 0.51× / 0.66× | 8206.5 / 24222.2 KiB |
| CSV short Unicode, file | CsvHelper | 1.31× / 0.96× | 323.3 / 529.0 KiB |
| CSV long notes, file | CsvHelper | 1.01× / 1.02× | 541.8 / 16530.8 KiB |
| CSV dense JSON, file | CsvHelper | 0.21× / 0.22× | 543.1 / 21069.9 KiB |
| CSV typed values, file | CsvHelper | 0.45× / 0.50× | 401.4 / 980.2 KiB |
| XLSX long plain, compact | SpreadCheetah | 1.53× / 2.94× | 245.7 / 224.0 KiB |
| XLSX long escaped, compact | SpreadCheetah | 2.28× / 2.12× | 245.5 / 227.3 KiB |
| XLSX long markup, compact | SpreadCheetah | 0.51× / 0.48× | 436.5 / 440.5 KiB |

These are specific equivalent export contracts. They do not support an overall library ranking. Short JSON is faster with CsvHelper; long plain and escaped XLSX are faster with SpreadCheetah. Dense quote runs and markup favor OfficeIMO, while long-note file timing is effectively tied despite a large allocation difference. Native and rotated runs use different invocation/hosting protocols, so compare versions or libraries within one protocol instead of comparing their absolute numbers across tables.

Full native reports: [CSV text](excel-csv-buffering-2026-09-08/native-csv-text.md), [CSV files](excel-csv-buffering-2026-09-08/native-csv-files.md), [compact XLSX](excel-csv-buffering-2026-09-08/native-excel-text.md), [public-default XLSX](excel-csv-buffering-2026-09-08/native-excel-defaults-final.md), and [reader/short-JSON comparison](excel-csv-buffering-2026-09-08/native-readers-final.md). [Native samples and allocations](excel-csv-buffering-2026-09-08/native-results.json) retain every measured case.

## Output and compatibility checks

- Eight before/after XLSX profiles have identical uncompressed package parts, including cell XML, styles, shared strings where applicable, and relationships. [Package evidence](excel-csv-buffering-2026-09-08/package-parts.json) records each part's hash and compressed package size. Compression boundaries can change total ZIP size without changing content. Long plain output grows from 59,419 to 60,693 bytes in compact mode (+2.14%) and from 68,720 to 70,283 bytes with public defaults (+2.27%); the other six profiles are unchanged or grow by less than 0.1%.
- All twelve CSV file shapes/quoting combinations produce identical UTF-8 bytes in both libraries and pass decoded-field validation. [File sizes](excel-csv-buffering-2026-09-08/file-output-sizes.json) include short Unicode, multi-character and Unicode delimiters, long notes, dense quotes, dates, decimals, booleans, and null strings.
- CSV tests pass on .NET 10 and .NET 8 (494 each), and .NET Framework 4.7.2 (366). The affected Excel matrix passes on .NET 8 (624, one skipped) and .NET Framework 4.7.2 (617, one skipped). Both libraries build for netstandard2.0; both benchmark projects build for .NET 8 and .NET 10.
- The full .NET 10 Excel run passes 3,897 tests, with five skipped. A desktop Excel COM smoke test fails twice while closing a pivot workbook, then passes with the final libraries in a separate probe and the original test location. The [pivot packages](excel-csv-buffering-2026-09-08/com-package-equivalence.json) match the passing baseline apart from generated relationship IDs. Two image-rendering baseline failures reproduced before these changes remain excluded; the run is not presented as an uninterrupted green full-suite result.
- Tests with hardware intrinsics disabled pass for CSV escaping (six cases) and Excel escaping/encoding (eleven cases). New tests cover buffer boundaries, split surrogates, strict encoding failure, disposal, pooled ownership, metadata length limits, cancellation, and package-name fallback.
- Public API reflection comparisons match for both libraries. One independent read-only review and a targeted confirmation report no actionable findings. The later validator-only code-page registration fix passes all twelve native default-profile cases.

The [validation record](excel-csv-buffering-2026-09-08/validation.json) retains counts and COM failures alongside successful rechecks. Remaining performance work stays in [the product roadmap](../ROADMAP.md#document-format-depth).
## Reproduction

Use the committed benchmark classes for native BenchmarkDotNet comparisons. The accompanying packet includes the isolated assembly loader, PowerForge driver, selected samples, and assembly hashes needed to reproduce the before/after run. Topology masks are specific to this workstation; discover and adapt them elsewhere. Build and validate first, then freeze the whole source checkout during PowerForge measurements because its provenance guard checks changes beyond the measured DLLs.
