# OfficeIMO MHTML comparison benchmarks

This opt-in project compares complete MHTML archive read and write workflows.
The read lane uses the same MimeKit-produced input in both implementations.
OfficeIMO performs its normal bounded MIME and HTML parsing; the comparison
uses MimeKit for MIME, AngleSharp for the HTML DOM, and retains every decoded
resource. The write lane serializes equivalent prepared multipart/related
models.

Preflight cross-reads both outputs and requires equal subject, root Content-ID,
root Content-Location, HTML, element count, ordered resource metadata, decoded
lengths, and payload hashes before timing.

```powershell
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Mhtml.Benchmarks.Comparisons -- validate
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Mhtml.Benchmarks.Comparisons -- --filter '*Mhtml*' --job Short --noOverwrite
dotnet run -c Release -f net10.0 --project .\OfficeIMO.Mhtml.Benchmarks.Comparisons -- evidence --repeat 3 --json .\.benchmark-artifacts\mhtml\evidence.json
```

The isolated runner records elapsed time and managed allocation per operation,
retained managed heap, sampled managed-heap growth, absolute process peak,
input and output bytes, decoded resource bytes, source commit, dirty-tree state,
runtime, and operating system.

The project stays outside `OfficeIMO.sln`, keeping MimeKit and benchmark-only
comparison dependencies out of normal restore, runtime, and package graphs.
