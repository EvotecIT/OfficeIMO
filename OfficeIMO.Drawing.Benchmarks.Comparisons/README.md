# OfficeIMO image library comparisons

This opt-in project compares equivalent image work without adding third-party dependencies to the OfficeIMO product or normal solution. The initial lanes cover full PNG and photographic JPEG decode, lossless PNG encode from one identical RGBA buffer, and 2048x1619-to-800x632 linear/bilinear resize from identical RGBA pixels.

Metadata identification remains in the first-party benchmark project, not the comparison table. OfficeIMO performs bounded container-structure validation during identification, while the candidate libraries' lightweight information APIs do not all validate the same structure; timing those unlike contracts as peers would be misleading.

The comparison uses SkiaSharp, Magick.NET, and StbImageSharp. ImageSharp 4 is intentionally excluded because its build now requires a Six Labors license key; benchmark restore/build must remain reproducible without secrets. StbImageSharp participates only in decode because it does not provide an encoder.

Correctness comes before timing. PNG decode and encode require exact full-buffer RGBA equality. The JPEG lane requires complete decode, equal dimensions, and a full-buffer mean absolute channel error no greater than 1.5 against every independent decoder so normal color-conversion and rounding differences remain distinguishable from material output defects.

```powershell
dotnet run --project OfficeIMO.Drawing.Benchmarks.Comparisons -c Release -f net10.0 -- --validate
dotnet run --project OfficeIMO.Drawing.Benchmarks.Comparisons -c Release -f net10.0 -- --job Dry --filter '*ImagePng*Benchmarks*'
```

BenchmarkDotNet's allocation column reports managed allocations only. SkiaSharp and Magick.NET perform material work in native code, so their managed allocation numbers are not total-memory comparisons; use process/native-memory profiling before drawing memory-efficiency conclusions across engines.
