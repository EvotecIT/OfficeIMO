# OfficeIMO.Email MIME comparisons

This opt-in BenchmarkDotNet project compares complete EML read and write
workflows through OfficeIMO.Email and MimeKit. It stays outside `OfficeIMO.sln`
so MimeKit does not enter OfficeIMO runtime restore, build, or packages.

The read lane gives both implementations the same MimeKit-generated EML and
includes parsing, body access, decoding every attachment, and consuming every
decoded payload byte. The write lane starts from equivalent prepared models
and serializes the complete message to a new in-memory byte array.

Validation runs before timing. It opens the shared input and both generated
outputs through both libraries and requires matching subject, sender, ordered To
and Cc recipients, normalized text and HTML bodies, ordered attachment names,
decoded lengths, and SHA-256 payload hashes. Output byte counts are comparable
only after those checks pass.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Email.Benchmarks.Comparisons -- validate --json .benchmark-artifacts\email\validation.json
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Email.Benchmarks.Comparisons -- --filter '*EmailMimeReadBenchmarks*' --job Short
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Email.Benchmarks.Comparisons -- --filter '*EmailMimeWriteBenchmarks*' --job Short
```

Use the shared library-comparison runner for provenance-backed evidence. Raw
BenchmarkDotNet output and machine-specific traces remain local.

Use the process-isolated evidence runner for retained managed heap, sampled
managed-heap peak, process working-set peak, and output size alongside elapsed
time and allocation. Read probes retain equivalent decoded bodies and
attachment payloads; write probes retain each complete serialized byte array.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Email.Benchmarks.Comparisons -- --evidence --repeat 3 --json .benchmark-artifacts\email\memory-evidence.json
```

The runner reports OfficeIMO divided by MimeKit. At or below `2.00x` is the
contender ceiling for every applicable dimension, not an optimality claim.
