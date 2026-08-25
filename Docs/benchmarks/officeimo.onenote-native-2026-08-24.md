# OneNote native workflow evidence — 2026-08-24

This run establishes native performance evidence for OfficeIMO.OneNote and
records the writer allocation work at commit
`fdca7dbdcef03ca3d4952346e8fb6f06a715986c`. It does not claim competitor
parity. No other managed library was identified that offers the same offline,
semantic desktop `.one` read/write contract.

## Contract and method

- Windows 11 `10.0.26200`, x64, .NET 10.0.11, 32 logical processors.
- The deterministic corpus contains 1, 25, or 100 pages with eight styled
  paragraphs per page.
- Create/write and read/edit/write disable the public writer's built-in
  round-trip validation inside the measured window. Each returned file is then
  reopened outside that window and checked against an exact SHA-256 fingerprint
  of ordered section, page, paragraph, run-text, and run-style fields.
- Each workflow sample runs in a fresh child process. It records elapsed time,
  managed allocation, sampled managed-heap growth, process peak, input bytes,
  output bytes, and provenance.
- BenchmarkDotNet separately measures the public default validated writer and
  serialization-only writer. The full run used 98-100 measurement iterations;
  its timing distributions were multimodal, so medians and allocation are more
  useful than single cold-process samples.

## Writer change

The writer now gathers reference identifiers in one traversal, builds the
global-ID map without quadratic list membership checks, and writes file-node
lists directly into the final package buffer instead of allocating and copying
an intermediate byte array for every list. A measured 128-byte initial buffer
avoids the repeated growth seen with 64 bytes without restoring the excess
capacity of 256 bytes.

Back-to-back 100-page create/write probes compared the preceding writer at
`1b137fccc57434812894bd623ce95e4598cf0313` with the selected implementation.
Five baseline probes had a 69.89 ms median and allocated 16.77 MiB. Five
confirmation probes had a 66.19 ms median and allocated 14.04 MiB. The final
clean three-sample run had a 65.42 ms median, 14.04 MiB allocated, 14.10 MiB
sampled managed-heap growth, 59.19 MiB median process peak, and a 482.51 KiB
output. That is about 6.4% faster and 16.3% lower allocation than the immediate
baseline, with identical output size and semantic fingerprint.

Allocation changes across the corrected isolated workflow are:

| Workflow | Scale | Before | After | Change |
| --- | --- | ---: | ---: | ---: |
| Create/write | Small | 215.70 KiB | 182.66 KiB | -15.3% |
| Create/write | Normal | 4.24 MiB | 3.55 MiB | -16.4% |
| Create/write | Large | 16.77 MiB | 14.04 MiB | -16.3% |
| Read/edit/write | Small | 695.67 KiB | 635.13 KiB | -8.7% |
| Read/edit/write | Normal | 10.89 MiB | 10.17 MiB | -6.6% |
| Read/edit/write | Large | 42.87 MiB | 40.11 MiB | -6.4% |

Read allocation is unchanged, as expected: the production change is confined
to serialization. Output sizes are also unchanged: 6,736 / 124,784 / 494,088
bytes for create/write and 11,792 / 129,840 / 499,136 bytes for
read/edit/write.

## Stable writer baseline

The full committed BenchmarkDotNet run on the 1-page and 25-page corpus records:

| Writer contract | Pages | Mean | Median | Allocated |
| --- | ---: | ---: | ---: | ---: |
| Default validated writer | 1 | 78.55 us | 75.88 us | 326.10 KiB |
| Serialization only | 1 | 41.48 us | 41.05 us | 151.83 KiB |
| Default validated writer | 25 | 2.511 ms | 2.544 ms | 7.01 MiB |
| Serialization only | 25 | 1.026 ms | 1.003 ms | 3.27 MiB |

The default writer intentionally reopens and semantically validates its output.
That safety cost is visible rather than being mislabeled as serialization cost.

## Regression gates and validation

`onenote-performance-budgets.json` contains ceilings for all nine workflow and
scale combinations. Allocation, sampled managed-heap growth, process peak, and
output bytes are hard gates. Fresh-process elapsed ceilings catch gross stalls;
BenchmarkDotNet and a same-machine healthy commit remain the timing authority.

The budget runner passed on .NET 8 and .NET 10. The OneNote product suite passed
306 tests on each of .NET 8 and .NET 10 and 304 tests on .NET Framework 4.7.2.
All file-producing evidence lanes reopened successfully with exact ordered
semantic fingerprints.
