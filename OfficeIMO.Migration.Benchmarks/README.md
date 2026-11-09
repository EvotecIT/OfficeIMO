# Migration runtime benchmarks

This opt-in runner measures existing migration contracts in Word 6/95, BIFF4,
iWork 2009, DBX and WordML 2003. It stays outside the normal solution and correctness
checks. PowerForge owns warmups, rotated sample ordering, timing, allocation,
optional process-memory sampling and JSON/CSV/Markdown reports.

| Profile | Small / medium / large | Shared input contract |
| --- | --- | --- |
| `Word69` | 100 / 1,000 / 10,000 runs | One ASCII paragraph, alternating bold runs; each FKP shares two character-property records |
| `Word69Notes` | 10 / 100 / 1,000 notes | FIB101; alternating automatic footnotes and endnotes, with two paragraphs per note |
| `Word95Notes` | 10 / 100 / 1,000 notes | FIB104; the same complete note contract as `Word69Notes` |
| `Biff4` | 1,000 / 4,000 / 16,000 rows | BIFF4 workspace with ordered cell-style findings and repeated global number-format findings |
| `IWork09` | 6 / 64 / 512 images | Pages 2009 XML with repeated references to one raster asset |
| `Dbx` | 6 / 64 / 512 messages | Indexed messages with distinct bodies, identities and attachment bytes |
| `OfficeXml` | 100 / 1,000 / 10,000 paragraphs | WordML 2003 with repeated paragraph and run property groups |

The note profiles require an explicit `-BenchmarkProfile` selection. The default
matrix retains the five original profiles.

The [2026-11-07 two-domain comparison](../Docs/benchmarks/officeimo.migration-runtime-2026-11-07.md) records all 1,848 validated samples, exact source/binary identities, allocation gains, slower observations and launch variation.

The [2026-11-08 Word notes comparison](../Docs/benchmarks/word69-notes-2026-11-08.md)
records 2,160 validated samples for the unchanged Word69 control and both note
profiles. Allocation and median budgets pass; three P95
checks fail. The report retains the complete matrix, raw numerical
observations and exact reconstruction of the note-preserving reference.

Input generation and complete output validation occur outside timing. The inputs
are synthetic scale workloads. Independent producer fixtures belong to the format
correctness evidence and do not establish a timing comparison when one revision
cannot preserve their content.

The in-process `net10.0` workload requires PowerShell 7.6 or later running on
.NET 10 or later, with PSPublishModule installed; the SDK selected for the build
does not select the PowerShell runtime.

Build with the repository-selected SDK:

```powershell
dotnet build OfficeIMO.Migration.Benchmarks/OfficeIMO.Migration.Benchmarks.csproj -c Release
```

Run a small smoke case before the full matrix:

```powershell
./Build/Benchmarks/Run-MigrationBenchmarks.ps1 -BenchmarkProfile Word69 -Scale Small -Plan
./Build/Benchmarks/Run-MigrationBenchmarks.ps1 -BenchmarkProfile Word69 -Scale Small -WarmupCount 1 -IterationCount 3
./Build/Benchmarks/Run-MigrationBenchmarks.ps1
./Build/Benchmarks/Run-MigrationBenchmarks.ps1 -BenchmarkProfile Dbx -Operation SearchContent
./Build/Benchmarks/Run-MigrationBenchmarks.ps1 -BenchmarkProfile Word69Notes,Word95Notes -Scale Small -WarmupCount 1 -IterationCount 3
```

`-BinaryRoot` selects a directory containing the workload host and the exact product
assemblies to measure. `-ProductSource` records the source commit or source-binding digest for an externally prepared binary root. Run each revision in a fresh PowerShell process so assembly
binding cannot mix revisions. Use the same workload source, input hashes, selected
cases, affinity, warmup/iteration policy and runtime for each revision. On hosts
with multiple cache or frequency domains, pass the discovered native-width mask
through `-ProcessorAffinityMask` and keep each domain in a separate result cohort.

`LoadInspect` includes public loading plus inspection or complete email-store
materialization. `ConvertSave` includes loading and normal DOCX, XLSX or EML output.
DBX's additional `SearchContent` operation eagerly searches every message body;
it is separate from `ReadAll`, which already has a bulk-read boundary.

BIFF4 conversion explicitly accepts the expected loss diagnostics, then verifies
every saved numeric value and finding location. The legacy iWork conversion uses
partial editable reconstruction for its existing content-limited diagnostic and
verifies every selected image. These policies are fixed across revisions. The
other cases require their declared loss-free contract. Every saved Open XML
artifact is schema validated and reopened; every EML is parsed again and checked.
A failed semantic check fails the benchmark lane.

The note profiles verify every body reference and its position, both note kinds,
complete ordered paragraph text, numbering and placement. Saved DOCX validation
also checks the note parts, reference identities, schema and public reopen. Their
semantic digest is separate from the diagnostic digest. A revision that omits
notes fails this contract and cannot serve as an equivalent performance baseline;
compare note optimizations with a fully correct feature baseline and retain the
original `Word69` profile as an unchanged control.

The executable also validates one input without measurement (including eager search for DBX) and prints its source
and diagnostic digests:

```powershell
dotnet OfficeIMO.Migration.Benchmarks/bin/Release/net10.0/OfficeIMO.Migration.Benchmarks.dll --validate OfficeXml 100
dotnet OfficeIMO.Migration.Benchmarks/bin/Release/net10.0/OfficeIMO.Migration.Benchmarks.dll --validate Word95Notes 1000
```

Inspect all raw samples, including regressions and repeated-launch variation.
Elapsed time and managed allocations describe the measured operation and include
the PowerShell invocation boundary. `-MemorySamplingIntervalMilliseconds 5` records
observed process-memory maxima as lower bounds in a separate instrumented run.
`-MeasureRetainedMemory` records a collected, process-wide managed-heap observation
after validation and result release; it includes host and cache effects. Neither
observation isolates native allocations or proves an unsampled peak.
