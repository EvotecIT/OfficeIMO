# OfficeIMO.Zip benchmarks

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

Before timing, validation requires exact agreement on ordered paths, names,
directory flags, depth, uncompressed lengths, UTC timestamps, total bytes, and a
SHA-256 structural fingerprint. The corpus contains 24, 512, and 4,000 files in
reverse package order, plus excluded directory entries, so deterministic sorting
is exercised by both implementations.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Zip.Benchmarks -- validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Zip.Benchmarks -- --filter '*ZipTraversalComparisonBenchmarks*'
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Zip.Benchmarks -- --evidence --repeat 3 --json .benchmark-artifacts\zip\evidence.json
```

BenchmarkDotNet reports elapsed time and managed allocations. The isolated
evidence runner additionally records input archive size and sampled peak managed
heap growth. ZIP creation is corpus setup rather than a measured OfficeIMO
product workflow, because `OfficeIMO.Zip` owns traversal policy rather than ZIP
writing.

The project stays outside `OfficeIMO.sln` so performance tooling does not affect
normal restore, build, test, or package dependencies.
