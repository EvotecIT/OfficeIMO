# OfficeIMO.Archives ZIP benchmarks

This opt-in benchmark project measures OfficeIMO's guarded, deterministic ZIP
metadata traversal against direct `System.IO.Compression` traversal of the same
safe archive. Both lanes open the same in-memory package, sort entries ordinally,
exclude directory entries, and materialize the same descriptor fields.

The platform lane deliberately omits OfficeIMO's path-traversal, depth,
entry-count, total-size, per-entry-size, and expansion-ratio policy. Ratios are
therefore the cost of that additional safety contract, not a claim that the raw
platform API provides equivalent protections. This workload remains an opt-in
policy-overhead diagnostic and is excluded from the published library-comparison
catalog.

On `net10.0` a OneArchive 1.2.0 lane (the package ships `net10.0` only) lists
the same archive through `Archive.Open` with `Format = ArchiveFormat.Zip` and
`EnumerateEntries()`, keeping regular files and sorting them ordinally. It runs
with OneArchive's default `ArchiveLimits` (entry count, declared total size,
path length and depth) and central-directory consistency checks, which apply
while reading. OneArchive's path-traversal rules apply only when extracting,
so this lane does not filter unsafe names the way OfficeIMO does. ZIP's DOS
timestamps carry no time zone: OneArchive reads them as UTC, while OfficeIMO and
`System.IO.Compression` read them as local time. The OneArchive lane reads the
stored DOS value as local time when the entry has no extended or NTFS
timestamp record, so all lanes compare the same stored field.

Before timing, validation requires exact agreement on ordered paths, names,
directory flags, depth, uncompressed lengths, UTC timestamps, total bytes, and a
SHA-256 structural fingerprint. The corpus contains 24, 512, and 4,000 files in
reverse package order, plus excluded directory entries, so deterministic sorting
is exercised by both implementations.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Archives.Benchmarks -- validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Archives.Benchmarks -- --filter '*ZipTraversalComparisonBenchmarks*'
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Archives.Benchmarks -- --evidence --repeat 3 --json .benchmark-artifacts\zip\evidence.json
```

BenchmarkDotNet reports elapsed time and managed allocations. The isolated
evidence runner additionally records input archive size and sampled peak managed
heap growth. ZIP creation is corpus setup rather than a measured OfficeIMO
product workflow, because `OfficeIMO.Archives` owns traversal policy rather than ZIP
writing.

The project stays outside `OfficeIMO.slnx` so performance tooling does not affect
normal restore, build, test, or package dependencies.

## Verified member copying

`VerifiedZipCopy` compares `OpenRead` followed by SHA-256 hashing with `CopyTo`
into a hashing destination. Both inspect the same 16 MiB Deflate member and
validate the complete result against its original SHA-256 in setup. Creation is
outside measurement. `CopyTo` includes private decoded-spool I/O; `OpenRead`
retains the complete decoded member array. This is an allocation and elapsed-time
tradeoff, not a peak-RSS bound or a throughput guarantee.

```sh
dotnet run -c Release -f net10.0 --project OfficeIMO.Archives.Benchmarks -- --filter '*VerifiedZipCopy*' --artifacts "$scratch/verified-copy"
```

Use a task-owned scratch directory for output and `TMPDIR`/`TEMP` for spool files.
The short job uses one launch, one warmup and three measured iterations; results
are workload- and host-specific. The [qualification record](../Build/Archives/README.md)
links the retained sample evidence.
