# OfficeIMO.Email performance evidence

The email performance contracts guard two common failure modes: copying a large MIME message too many times and scaling mailbox parsing by message count rather than source size.

Run the evidence with:

```powershell
dotnet test OfficeIMO.Email.Tests/OfficeIMO.Email.Tests.csproj `
    -c Release `
    -f net8.0 `
    --filter FullyQualifiedName~EmailPerformanceEvidenceTests `
    --logger "console;verbosity=detailed"
```

The tests measure allocations on the parsing thread after constructing the fixture. They also apply a generous time ceiling to catch hangs and accidental super-linear work without turning ordinary machine variance into failures.

## Current baseline

Measured on 2026-07-10 with an Apple M4, 24 GB memory, macOS 26.5, .NET 8.0.23 runtime, and .NET SDK 10.0.102:

| Workload | Source bytes | Parsing-thread allocations | Elapsed |
| --- | ---: | ---: | ---: |
| One 1 MiB decoded plain-text MIME body | 1,435,016 | 13,123,064 | 3.4 ms |
| 500-message mbox archive | 120,850 | 4,554,840 | 12.0 ms |

These numbers are a local regression baseline, not a cross-machine throughput promise. The committed contracts enforce allocation ceilings proportional to source size plus fixed headroom and a ten-second hang ceiling. Returned strings, message models, and requested attachment payloads are intentionally included in the allocation measurement.

For realistic deployments, measure representative `.msg`, `.eml`, TNEF, and mbox corpora with the same reader options used by the application. In particular, `includeAttachmentContent: false` changes retained memory materially when callers only need metadata.
