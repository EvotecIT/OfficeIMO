# Wide quote-free CSV field-span read — 2026-09-28

`CsvDocument.ReadFieldSpansFromText` used the text-data-reader batch parser for
large strings even when the input contained no quotes. A focused CPU-sampling
trace of the validated 25,000-row, 40-column workload placed 97% of sampled
time inside that batch path. The trace identifies the executed path; its sample
percentages are not elapsed-time measurements.

Commit `59d0aa6b1` routes strings of at least 64 KiB with no `"` through the
existing direct field-span parser. It forwards the quote-free observation so
that parser does not scan the string a second time. Quoted input and shorter
strings retain the prior dispatch. The direct parser retains the option
fallbacks, projection, skipped-record handling, and cancellation checks.

The comparison below used the same semantically validated 25,000 × 40 input
for both implementations. Each reader visited every field and returned the
expected checksum before timing. Baseline and candidate were run in the order
B–C–B–C; each run included Sep as an environment control. The baseline was
`dd5b58d51`, and the candidate source is `59d0aa6b1`. BenchmarkDotNet 0.15.8
ran on Windows 11 build 26200.9457, AMD Ryzen 9 9950X3D2, .NET 10.0.12,
AboveNormal priority, four warmups, eight measured iterations, one launch,
and retained outliers. The two affinity masks select separate groups of 16
logical processors.

| Run | Affinity | OfficeIMO mean | Sep mean | OfficeIMO allocation |
| --- | --- | ---: | ---: | ---: |
| Baseline 1 | `0xFFFF` | 4.685 ms | 3.044 ms | 680 B |
| Candidate 1 | `0xFFFF` | 2.315 ms | 2.481 ms | 456 B |
| Baseline 2 | `0xFFFF` | 4.189 ms | 2.783 ms | 680 B |
| Candidate 2 | `0xFFFF` | 2.888 ms | 3.040 ms | 456 B |
| Baseline 1 | `0xFFFF0000` | 4.095 ms | 2.706 ms | 680 B |
| Candidate 1 | `0xFFFF0000` | 2.766 ms | 2.830 ms | 456 B |
| Baseline 2 | `0xFFFF0000` | 5.842 ms | 2.560 ms | 680 B |
| Candidate 2 | `0xFFFF0000` | 2.217 ms | 2.555 ms | 456 B |

The candidate was faster in both rotated pairs on both processor groups and
allocated 224 fewer managed bytes per operation. Its difference from Sep is
small relative to the variation between runs, so this evidence supports a
source-level improvement, not a general library ranking. The complete CSV
suite passed on .NET 8 and .NET 10 (603 tests each) and .NET Framework 4.7.2
(430 tests); the CSV library built on all four target frameworks, including
`netstandard2.0`. The .NET 8 suite also passed 603 tests under Linux/WSL after
two delimiter and encoding tests were corrected to expect the documented
platform-default newline. The tests include long, quote-free, 40-column CRLF input,
Unicode characters whose low byte resembles CSV punctuation, projected last
fields, a final row without a newline, and cancellation between rows.

The refreshed 14-method .NET 8.0.31 snapshot on `0xFFFF` measured OfficeIMO's
same span visitor at 2.441 ms / 456 B, Sep at 2.971 ms / 5,432 B, and Sylvan
at 3.806 ms / 49,657 B. It validated all selected read and write outputs and
generated the [current compact comparison](readme-current/officeimo.csv.comparison.json).
Absolute write times for several unrelated libraries also rose from the earlier
same-day snapshot; this run alone does not establish a write regression. The
rotated .NET 10 results above are the before/after evidence for the parser
change.

To measure the current implementation under the same two-domain policy, run:

```powershell
$env:OFFICEIMO_CSV_WIDE_ROW_COUNT = '25000'
dotnet run -c Release -f net10.0 --project OfficeIMO.CSV.Benchmarks -- --filter '*CsvWideBenchmarks.OfficeIMO_ReadTextFieldSpanVisitorSkipHeader' '*CsvWideBenchmarks.Sep_ReadFieldSpans' --affinityMasks 0xFFFF,0xFFFF0000 --priority AboveNormal --warmupCount 4 --iterationCount 8 --launchCount 1 --outliers DontRemove
Remove-Item Env:OFFICEIMO_CSV_WIDE_ROW_COUNT
```

This is an in-memory span visitor measurement. It does not establish file-I/O,
peak-memory, quoted-input, other row-shape, or native Linux/macOS performance
budgets. Those contracts remain separate in the [roadmap](../ROADMAP.md).
