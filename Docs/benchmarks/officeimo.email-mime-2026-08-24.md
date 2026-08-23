# Email MIME comparison evidence

This 2026-08-24 run compares complete EML read and write workflows through
OfficeIMO.Email and MimeKit 4.17.0. The read lane parses the same EML, accesses
both bodies, decodes every attachment, and consumes every decoded byte. The
write lane serializes equivalent prepared models to new in-memory byte arrays.

Validation opens the shared input and both outputs through both libraries and
requires equal subject, sender, ordered To and Cc recipients, normalized text
and HTML bodies, ordered attachment names, decoded lengths, and SHA-256 payload
hashes before timing or size evidence is accepted.

This was a full BenchmarkDotNet run from clean source commit
`8cff127ec2f4e4e28b1bcc6013a4910d25570a45` on Windows 11
10.0.26200.9168, an AMD Ryzen 9 9950X3D2 with 16 physical cores, .NET SDK
10.0.111, and .NET 10.0.11 x64. Allocation is managed allocation per operation.
The host produced multimodal timing distributions, so the table includes both
the mean and median; the contender classification is unchanged by either.

| Workflow | Scale | Implementation | Mean | Median | Allocation | Output bytes | OfficeIMO time ratio | OfficeIMO allocation ratio | OfficeIMO size ratio |
| --- | --- | --- | ---: | ---: | ---: | ---: | ---: | ---: | ---: |
| Read | Small | OfficeIMO | 20.80 us | 20.15 us | 37.20 KB |  | 1.02x | 0.58x |  |
| Read | Small | MimeKit | 20.30 us | 20.09 us | 63.68 KB |  |  |  |  |
| Read | Normal | OfficeIMO | 413.12 us | 400.15 us | 296.53 KB |  | 1.12x | 0.37x |  |
| Read | Normal | MimeKit | 369.86 us | 367.13 us | 812.03 KB |  |  |  |  |
| Write | Small | OfficeIMO | 10.13 us | 10.29 us | 63.73 KB | 8,528 | 1.29x | 1.38x | 1.08x |
| Write | Small | MimeKit | 7.85 us | 7.69 us | 46.16 KB | 7,868 |  |  |  |
| Write | Normal | OfficeIMO | 219.59 us | 215.37 us | 1,351.70 KB | 240,424 | 1.30x | 1.52x | 1.07x |
| Write | Normal | MimeKit | 168.54 us | 161.71 us | 890.96 KB | 224,479 |  |  |  |

Every lane is within 2x for both time and allocation. OfficeIMO read allocation
is materially lower than MimeKit. The buffered byte-array base64 path used by
ordinary in-memory bodies and attachments reduced the measured normal write
from about 2.85x to 1.30x time while keeping source-backed and payloads above
1 MiB on the bounded streaming path.

Reproduce the validation and full runs with:

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Email.Benchmarks.Comparisons -- validate --json .benchmark-artifacts\email\validation.json
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload emailmimeread -RunMode full -Framework net10.0
.\Build\Run-LibraryComparisonBenchmarks.ps1 -Workload emailmimewrite -RunMode full -Framework net10.0
```

The comparison project remains outside `OfficeIMO.sln`; MimeKit is an opt-in
benchmark dependency and does not enter OfficeIMO runtime restore or packages.
Raw BenchmarkDotNet output remains local.

## PST scale regression

The existing `TwoThousandMessagePstUsesBoundedRetainedManagedMemory` contract
also exposed a separate writer regression during this pass. An isolated .NET 10
run took 68.19 seconds before the fix, exceeding its 45-second ceiling. Small
properties and single-block heaps were needlessly routed through disk-backed
data-tree journals. Commit `2e334d57e15de88c524554c14c94a02272bf64c2`
writes payloads at or below the 8,176-byte PST block limit directly and retains
the journaled path for larger data trees.

The committed-head rerun completed in 4.03 seconds, a 16.9x reduction. It wrote
2,000 items to a 4,080,640-byte PST, retained 733,072 managed bytes, reopened the
artifact, and enumerated all 2,000 items. Boundary coverage separately round
trips 8,176-byte direct and 8,177-byte journaled data trees.
