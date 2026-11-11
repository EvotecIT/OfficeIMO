# Email MIME comparison evidence

This 2026-11-11 ARM64 run compares complete EML read and write workflows with
MimeKit 5.0.0. OfficeIMO reads are 19–25% faster and writes are 13–17% faster in
the final application-policy cohort. Read allocation is 0.37–0.63x MimeKit;
write allocation is 0.36–0.68x. Every lane passes the existing semantic validator.
OfficeIMO output also matches the original implementation byte for byte.

## Sources and method

The baseline is `b43eb05be6970e2bced6d5d6a2b017c711f6943c`, which only moves
the comparison pin to MimeKit 5.0.0. The optimized source is
`13a350bf0d29e9046c34830d0ab56aeb4721b6d7`. Both use the unchanged comparison
corpus and cross-reader validator in `OfficeIMO.Email.Benchmarks.Comparisons`.
MimeKit remains an opt-in benchmark dependency outside `OfficeIMO.slnx` and
OfficeIMO runtime packages.

The host is an Apple M4 with 10 cores, macOS 27.0.1 (26A434), .NET SDK 10.0.404
and .NET 10.0.15 Arm64. BenchmarkDotNet 0.15.8 uses six warmups, twelve measured
iterations and two launches per case. No outliers are removed. Allocation is
managed allocation per operation; 1 KiB is 1,024 bytes. Ratios divide the
OfficeIMO mean by the MimeKit mean.

The final paired confirmation applies `taskpolicy -a -t 0 -l 0` equally to both
builds and clears the child controller's background policy with `taskpolicy -B`.
It runs the optimized cohort first, then the baseline, reversing the earlier
pair's order. The shared heavy-run lock serializes the complete jobs. This is
application scheduling policy, not core affinity or a quiet-host reservation.
BenchmarkDotNet cannot raise process priority on this host; both engines run
without that priority elevation.

The host has variable background activity, including compilation and OS indexing.
Earlier cohorts produced large timing spreads and some losses. They are retained
in the [machine-readable evidence](officeimo.email-mime-2026-11-11.json), alongside
all final paired result samples, medians, standard deviations and source hashes.
The final pair is substantially steadier. These are measured workloads and host
policies, not a ranking for every processor or scheduling condition.

The commands inside each held lock are a Release build with `-m:2 -nodeReuse:false`,
followed by:

```sh
dotnet OfficeIMO.Email.Benchmarks.Comparisons/bin/Release/net10.0/OfficeIMO.Email.Benchmarks.Comparisons.dll validate --json validation.json
dotnet OfficeIMO.Email.Benchmarks.Comparisons/bin/Release/net10.0/OfficeIMO.Email.Benchmarks.Comparisons.dll --filter '*EmailMime*' --warmupCount 6 --iterationCount 12 --launchCount 2 --outliers DontRemove --exporters json --artifacts results
```

Build subprocesses disable MSBuild server and node reuse. Validation opens the
shared input and both generated outputs through both libraries and checks
subject, sender, ordered To/Cc recipients, normalized text/HTML bodies, ordered
attachment names, decoded lengths and SHA-256 payload hashes before timing.

## Before: MimeKit 5.0 baseline

| Workflow | Scale | OfficeIMO mean ± SD (us) | MimeKit 5.0 mean ± SD (us) | OfficeIMO / MimeKit time | Allocation OfficeIMO / MimeKit (KiB) |
| --- | --- | ---: | ---: | ---: | ---: |
| Read | Small | 27.912 ± 0.041 | 16.801 ± 0.446 | 1.66x | 41.66 / 59.18 |
| Read | Normal | 668.014 ± 5.829 | 277.182 ± 2.270 | 2.41x | 304.13 / 808.19 |
| Write | Small | 8.014 ± 0.154 | 5.756 ± 0.072 | 1.39x | 66.14 / 33.71 |
| Write | Normal | 187.397 ± 5.032 | 115.274 ± 0.543 | 1.63x | 1354.27 / 827.31 |

## After: shared Core codec and MIME scanning

| Workflow | Scale | OfficeIMO mean ± SD (us) | MimeKit 5.0 mean ± SD (us) | OfficeIMO / MimeKit time | Allocation OfficeIMO / MimeKit (KiB) |
| --- | --- | ---: | ---: | ---: | ---: |
| Read | Small | 12.704 ± 0.084 | 16.913 ± 0.405 | 0.75x | 37.07 / 59.18 |
| Read | Normal | 224.803 ± 1.225 | 276.550 ± 1.756 | 0.81x | 296.71 / 808.19 |
| Write | Small | 5.068 ± 0.078 | 5.800 ± 0.072 | 0.87x | 23.09 / 33.71 |
| Write | Normal | 108.943 ± 5.867 | 131.812 ± 8.414 | 0.83x | 295.72 / 827.31 |

Compared with the original OfficeIMO implementation on the same host and policy,
read time falls by 55% for Small and 66% for Normal; write time falls by 37% and
42%. Allocation falls by 11% / 2% for reads and 65% / 78% for writes. Cross-cohort
percentages still include host variation; the tables retain each MimeKit control.

## Historical MimeKit 4.17 reference

The [2026-08-24 evidence](officeimo.email-mime-2026-08-24.md) remains unchanged.
It ran on Windows 11, a Ryzen 9 9950X3D2 and .NET 10.0.11 x64. These numbers
provide historical context, not a direct version comparison with this Mac.

| Workflow | Scale | OfficeIMO / MimeKit 4.17 mean (us) | Allocation OfficeIMO / MimeKit (KiB) |
| --- | --- | ---: | ---: |
| Read | Small | 20.80 / 20.30 | 37.20 / 63.68 |
| Read | Normal | 413.12 / 369.86 | 296.53 / 812.03 |
| Write | Small | 10.13 / 7.85 | 63.73 / 46.16 |
| Write | Normal | 219.59 / 168.54 | 1351.70 / 890.96 |

## What changed

The original Normal-read trace spends most sampled managed thread time in scalar
base64 length validation and whitespace compaction. Core now owns MIME analysis,
whitespace compaction, bounded decoding and byte-oriented encoding. Net8/net10
use accelerated platform primitives and vectorized searches; older targets retain
scalar kernels. Email reuses the analysis for decoding, searches bounded blocks
for line-start boundary candidates, then applies the unchanged full delimiter
predicate. Quoted-printable decoding copies literal runs in bulk and retains
malformed diagnostics and limit checks.

The original write trace also exposes GC, output-buffer copies and deterministic
boundary hashing. Email reuses the common boundary hash and per-operation UTF-8
body bytes. Core encodes into byte buffers; Email folds backwards into the same
storage. A sequential pooled Core output stream avoids a growing temporary result
array for modern generated-EML writes. The existing bounded writer still owns
quotas and accounting. Boundary hashing remains content-dependent and serial;
changing its hash would change serialized bytes.

Four Core image/SVG/font data-URI paths decode string suffixes without allocating
another large string. The focused 4 KiB comparison drops from 6.967 to 3.608 us
and 14.72 to 4.02 KiB; 1 MiB drops from 1463.55 to 491.00 us and 3755.03 to
1024.17 KiB. These are separate component measurements. Other `Convert` call
sites retain the runtime's accelerated implementation. There is no mechanical
call-site rewrite, hand-written SIMD kernel or new runtime dependency.

Sampled thread-time traces include GC and waits; their percentages are not native
CPU instruction costs. The complete validated lanes determine the final result.

## Compatibility proof

- Full Email suites: 1,883 passed and 13 optional skips on each of net8/net10.
- Full Core/Drawing suites: 6,615 passed and 5 optional skips on each framework
  at `ef7635b914`; the later changes affect only Email boundary scanning and docs.
- Shared decoder consumers at `5d2b76e835`: Word 8,122 passed / 10 skipped,
  Excel 8,405 / 5, PDF 13,510 / 0. The suffix decoder is unchanged afterward.
- Core and Email netstandard2.0 build with zero warnings and errors.
- The same 32 codec vectors pass with hardware intrinsics disabled and against
  the actual netstandard2.0 Core assembly. The modern Drawing test project also
  uses modern-only APIs, so its whole graph is not used as the scalar harness.
- The actual scalar Email graph passes the transfer/boundary vectors. Seven
  boundary cases cover 64 KiB crossings, CR/LF/CRLF, invalid delimiter prefixes,
  closing whitespace, nonzero ranges, empty parts and truncated markers.
- Original/candidate snapshots match 224 malformed-input observations per target,
  including content, lengths, diagnostics, exceptions and processing budgets.
  Original/candidate Small and Normal EML outputs match exactly; SHA-256 values
  are retained in the JSON evidence.

Correctness runs exclude measurement-only `Category=Performance` cases. Optional
native/live fixtures retain their existing skips. Net472 requires Windows; the
local scalar proof here uses the actual netstandard2.0 libraries.

## Output size and limits

| Scale | Input/MimeKit output bytes | OfficeIMO output bytes | Decoded attachments |
| --- | ---: | ---: | ---: |
| Small | 7,760 | 8,528 | 1 / 4,096 bytes |
| Normal | 221,888 | 240,424 | 4 / 131,072 bytes |

OfficeIMO retains CRLF and base64 body parts. MimeKit uses LF on this Mac and
does not base64-encode the two text body parts in this fixture. Their default
serialization policies differ; validation establishes equal decoded content.
The 108-byte Small and 2,591-byte Normal differences from historical MimeKit
output equal the LF line counts, so those size changes do not establish a version
improvement. OfficeIMO bytes remain unchanged.

No lane is slower in the final application-policy cohort. Earlier ordinary-host
runs include timing losses and inconclusive spreads, especially Small reads and
Normal writes; the JSON retains them. The comparison measures these two complete
in-memory EML workloads, not every MIME shape or large source-backed stream.
Mail transport and authentication/signature protocols are outside this work.
