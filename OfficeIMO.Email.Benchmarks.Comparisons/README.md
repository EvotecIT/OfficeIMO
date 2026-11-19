# OfficeIMO.Email MIME comparisons

This opt-in BenchmarkDotNet project compares complete EML read and write
workflows through OfficeIMO.Email and MimeKit 5.0.0. It stays outside `OfficeIMO.slnx`
so MimeKit does not enter OfficeIMO runtime restore, build, or packages.

The 2026-11-11 evidence
records the validated MimeKit 5.0 baseline, optimized results and historical 4.17 reference.

The read lane gives both implementations the same MimeKit-generated EML and
includes parsing, body access, decoding every attachment, and consuming every
decoded payload byte. The write lane starts from equivalent prepared models
and serializes the complete message to a new in-memory byte array.

On `net10.0` the project also runs a OneMail 1.0.1 lane (the package ships
`net10.0` only). Its read lane parses the same input with `Mail.Parse` over the
in-memory bytes, reads both bodies through `ReadText()`, and copies every
decoded attachment from `IMailAttachment.Open()`. Its write lane starts from a
prepared `OutgoingMessage` and times `ToEntity()` plus `ToArray()`, so building
the MIME tree and encoding the parts are measured as they are for the other
lanes.

Validation runs before timing. It opens the shared input and every generated
output through every library and requires matching subject, sender, ordered To
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

The runner measures OneMail as well on `net10.0`, includes it in the
equivalence check, and reports ratios as OfficeIMO divided by MimeKit. At or below `2.00x` is the
contender ceiling for every applicable dimension, not an optimality claim.
