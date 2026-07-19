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

Measured on 2026-07-11 with an Apple M4, 24 GB memory, macOS 26.5, .NET 8.0.23 runtime, and .NET SDK 10.0.102:

| Workload | Source bytes | Parsing-thread allocations | Elapsed |
| --- | ---: | ---: | ---: |
| One 1 MiB decoded plain-text MIME body | 1,435,016 | 13,123,448 | 3.7 ms |
| One MSG with a 1 MiB attachment | 1,063,936 | 5,477,024 | 10.2 ms |
| 500-message mbox archive | 120,850 | 4,750,648 | 9.9 ms |

These numbers are a local regression baseline, not a cross-machine throughput promise. The committed contracts enforce allocation ceilings proportional to source size plus fixed headroom and a ten-second hang ceiling. Returned strings, message models, and requested attachment payloads are intentionally included in the allocation measurement.

The streaming path has a separate retained-memory contract: a generated 16 MiB attachment is written without a
whole-artifact buffer, reopened as file-backed content, and held under 8 MiB of additional retained managed memory.
All EML, MSG/OFT, and TNEF streaming tests also reject whole-payload source reads and destination writes.

For realistic deployments, measure representative `.msg`, `.eml`, TNEF, and mbox corpora with the same reader options used by the application. In particular, `includeAttachmentContent: false` changes retained memory materially when callers only need metadata.

## PST and OST large-store evidence

`OfficeIMO.Email.Store` uses deterministic I/O ceilings instead of a wall-clock assertion for its committed large
store contract. The synthetic source reports a 64 GiB length while containing a small valid PST at the offsets the
reader visits. One test opens the session, enumerates and summarizes an item, selectively reads its body and
attachment metadata, proves the attachment payload is untouched until the first stream read, searches body text,
and runs structural validation. The contract allows less than 4 MiB of total source reads and no individual read
larger than 128 KiB.

Run it with:

```powershell
dotnet test OfficeIMO.Email.Tests/OfficeIMO.Email.Tests.csproj `
    -c Release `
    -f net8.0 `
    --filter FullyQualifiedName~EmailStoreSessionTests
```

An aggregate-only validation on 2026-07-16 used a private 22,416,596,992-byte Outlook OST without exporting or
retaining messages, names, identifiers, snippets, or hashes. The observed run found 137 folders and 197,159
declared items, selectively read 50 items without read failures, streamed 29 attachment payloads into a bounded
sample, scanned two resumable 50-item content-search batches, read seven typed appointment items, and projected 20
items through Reader. The tracked store session read 27,501,637 source bytes. A separate bounded integrity pass
checked 500 B-tree pages and 1,000 blocks (2,910,483 structural bytes) in 127 ms with no structural failures; it
reported truncation because it stopped at the configured limits.

Those numbers prove behavior against one large real file, not general throughput. The repeatable guarantees are
the configured page, block, decoded-property, searchable-character, attachment, item, and source-read bounds.

The managed PST writer has cardinality-scale contracts independent of private mail: 2,000 deterministic messages
must remain under 64 MiB of retained managed growth, finish within 45 seconds, and reopen with the exact item count.
Separate 100,000-entry gates keep conversion mappings under 32 MiB and the semantic deduplication index under
16 MiB of retained growth. The latter two structures are sequential/disk-backed rather than retained item lists.

## Outlook OAB evidence

`OfficeIMO.Email.AddressBook` keeps only metadata and the active record while enumerating. Synthetic v4 fixtures
exercise dynamic property tables, every supported scalar and multi-valued encoding, raw-byte retention, exact-offset
search resume, CRC/framing/full-decode validation, cancellation, and configured limits on every target framework.

An aggregate-only run on 2026-07-16 inspected a private Outlook cache containing 18 OAB components and three v4
Full Details address lists. All 8,049 declared entries decoded, the object-type totals reconciled to the declared
count, all three seeded CRC values matched, and full framing/schema validation completed with zero skipped records
or session diagnostics. The combined open, decode, and second full validation pass completed in 386 ms on the
current Windows test machine. No names, addresses, identifiers, properties, record bytes, or hashes were printed,
stored, or copied into the repository.

This is compatibility evidence for one Outlook cache, not a throughput promise. The durable large-file contract is
the bounded discovery/schema/record model, sequential one-record memory behavior, exact-offset checkpoints, optional
raw-byte retention, and explicit checksum, record, string, binary, value-count, search, and Reader limits.
