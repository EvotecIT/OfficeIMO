# XLSX long-text export, 2026-09-28

The direct XLSX writer now returns from its scalar XML-escape loop when a dense
run of escapes is followed by 32 ordinary characters. It resumes the vector
search at the first unwritten character. This reduces the cost of long text
that begins with closely spaced XML escapes and then has a long plain suffix.

The baseline is source commit `4de45f7f7`, Excel assembly SHA-256
`7AE94598F3F8A00B22EB01171DA79677B8AF6E3D65E716E5DCE5340F754FD4A4`.
The candidate is the writer change documented here, Excel assembly SHA-256
`7258FEB12AEF084FBFD5765B5CDEFD0A29DCB45F857DD70E8472DDFDE2D26356`.
Both comparison directories contain the same benchmark and dependency binaries;
only `OfficeIMO.Excel.dll` differs. The source and setup follow the existing
[isolated comparison runner](excel-csv-buffering-2026-09-08/reproduction/README.md).

## Correctness

The `--xml-compare` mode reopened and validated 1,000-row public-writer output
for six shapes. Every corresponding uncompressed ZIP part had the same SHA-256
hash before and after, and each complete package had the same byte length.
The [part hashes and sizes](xlsx-long-text-2026-09-28/package-parts.jsonl)
include compact and public-default long plain and escaped output, compact dense
markup, and compact short plain output. A public `WriteDataReader` test also
reopens compact, default inline, and default shared-string output. It checks
the original text after XML-control sanitization, the carriage-return entity,
and Open XML package validation on .NET 8, .NET 10, and .NET Framework 4.7.2.

## Rotated Windows result

PowerForge measured eight validated exports per operation with 12 warmups and
24 retained measured samples per engine and scenario. Before/after operations
were rotated. The host was Windows 11 build 26200, .NET 10.0.12, Normal process
priority, on a Ryzen 9 9950X3D2. Processor masks `0xFFFF` and `0xFFFF0000`
select its two 16-logical-processor domains. The table gives after/before
ratios for the exact candidate assembly above; lower is faster.

| Public writer shape | Domain A mean | Domain A median | Domain B mean | Domain B median |
| --- | ---: | ---: | ---: | ---: |
| Compact long escaped | 0.42 | 0.47 | 0.54 | 0.60 |
| Default long escaped | 0.68 | 0.57 | 0.67 | 0.63 |
| Compact long plain | 0.99 | 1.00 | 1.02 | 1.00 |
| Default long plain | 1.01 | 0.99 | 0.97 | 1.07 |
| Compact long dense markup | 1.06 | 1.03 | 1.02 | 0.99 |
| Compact short plain | 0.85 | 1.03 | 1.07 | 1.04 |

The [Domain A](xlsx-long-text-2026-09-28/domain-a-summary.json) and
[Domain B](xlsx-long-text-2026-09-28/domain-b-summary.json) summaries and
their [A](xlsx-long-text-2026-09-28/domain-a-samples.csv) and
[B](xlsx-long-text-2026-09-28/domain-b-samples.csv) raw samples retain
outliers. Plain-text controls vary around parity. Dense markup shows a small
possible regression on Domain A. The shared workstation's load affects
absolute times, so the rotated escaped-text gain is stronger evidence than any
one absolute mean.

The old long-plain 25-30% reduction target is still open: this change does not
alter its escape-free path. Allocation, peak memory, native Linux and macOS,
and portable absolute budgets were not qualified in this run. The remaining
work is tracked in the [single roadmap](../ROADMAP.md).

## Long-plain follow-up profile

On source commit `c7a431094`, the same 1,000-row long-plain fixture allocated
251,287 bytes per compact export and 645,949 bytes per public-default export
in a warmed, single-thread allocation probe. A 10,000-export sampled trace on
Domain A placed `PooledUtf8TextWriter.FlushBytes` on 90.48% of compact and
75.52% of default sampled stacks. Native Deflate appeared on 45.48% and
43.50%, respectively; the default shared-string probe appeared on 10.86%.
These are overlapping inclusive stack percentages, not an additive breakdown
or elapsed-time budget. The compact trace gives little support for further
XML-scanner tuning as the route to the long-plain target.

A diagnostic 32 KiB writer buffer kept all six uncompressed package-part hashes
identical and changed four compressed package sizes by one byte. Its short
rotated Domain A smoke run made compact long-plain export 1.13 times the
baseline mean and default long-plain 1.01 times the baseline mean. That
candidate was reverted. This six-iteration smoke result does not establish a
portable regression; it only rules out accepting the smaller buffer without
stronger evidence. Reproduce the stage probe with the existing
[isolated comparison runner](excel-csv-buffering-2026-09-08/reproduction/README.md)
using `--profile ExcelLongPlain 10000 65535 candidate` or
`--profile ExcelDefaultLongPlain 10000 65535 candidate` and a sampled
`dotnet-trace` collection of the child process.

A separate 128 KiB byte-buffer diagnostic on source commit `60fa82a54`
preserved every uncompressed package-part hash and the complete package size
for compact and default long plain/escaped text, dense markup, and short plain
text. The baseline and candidate Excel assembly hashes were
`219A919690E3CC7FD8E7E8E1853BF860760971B9F78C8115F0287B30769D9D75`
and `1E4CEB0002E531F6B014DFE1C6DCAF6879A661587FC6326C0333605CD9EE98FD`.
In a Domain A smoke comparison with two warmups, four retained measurements,
four exports per operation, and rotated before/after order, the candidate's
mean-time ratios were 1.04 for compact long plain, 0.98 for default long plain,
and 1.08 for dense markup. This small run did not support the long-plain target;
the buffer change was reverted. It is not a full two-domain regression result.
The [part hashes](xlsx-long-text-2026-09-28/buffer-128-diagnostic/package-parts.jsonl),
[comparison](xlsx-long-text-2026-09-28/buffer-128-diagnostic/comparison.json),
and [raw samples](xlsx-long-text-2026-09-28/buffer-128-diagnostic/samples.json)
retain the diagnostic evidence.
