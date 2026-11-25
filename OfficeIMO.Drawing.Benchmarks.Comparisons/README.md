# OfficeIMO image library comparisons

This opt-in project compares equivalent image work without adding third-party dependencies to the OfficeIMO product or normal solution. The lanes cover PNG and photographic JPEG decode into a managed RGBA byte array, lossless PNG encode from one identical RGBA buffer, 2048x1619-to-800x632 linear/bilinear resize, static lossless VP8L WebP, and classic TIFF LZW. TIFF and WebP interchange is checked in both directions before timing; the TIFF validator also selects an independently written second page. Lossless WebP decode runs over four libwebp-encoded inputs: the synthetic pattern, the photo, a UI screenshot and the transparent logo, each written with `exact` so every decoder must return the source pixels.

Metadata identification remains in the first-party benchmark project, not the comparison table. OfficeIMO performs bounded container-structure validation during identification, while the candidate libraries' lightweight information APIs do not all validate the same structure; timing those unlike contracts as peers would be misleading.

The comparison uses SkiaSharp, Magick.NET, and StbImageSharp. ImageSharp 4 is intentionally excluded because its build now requires a Six Labors license key; benchmark restore/build must remain reproducible without secrets. StbImageSharp participates only in decode because it does not provide an encoder.

OneImage (pure managed) runs in every lane on `net10.0` only, because the package targets only that framework. It decodes with EXIF auto-orientation off, so every decoder returns the stored layout; encodes PNG at its default deflate level 6, TIFF LZW with the horizontal predictor, and lossless WebP at its default effort 4; and resizes with its span-to-span `PixelResizer`, triangle filter, on the stored sRGB values. Its output passes the same checks as the other engines, and in the TIFF and WebP lanes OneImage also has to decode the OfficeIMO and Magick.NET files exactly.

Correctness comes before timing. PNG, TIFF LZW, and lossless WebP require exact full-buffer RGBA equality. The JPEG lane requires complete decode, equal dimensions, and a full-buffer mean absolute channel error no greater than 1.5 against every independent decoder so normal color-conversion and rounding differences remain distinguishable from material output defects.

```powershell
dotnet run --project OfficeIMO.Drawing.Benchmarks.Comparisons -c Release -f net10.0 -- --validate
dotnet run --project OfficeIMO.Drawing.Benchmarks.Comparisons -c Release -f net10.0 -- --job Dry --filter '*ImagePng*Benchmarks*'
```

On heterogeneous or multi-CCD processors, collect comparative evidence through the repository runner with a reviewed affinity mask that keeps the benchmark on one representative processor region. For example, this host exposes one CCD as logical processors 0-15:

```powershell
./Build/Run-LibraryComparisonBenchmarks.ps1 -Profile full -Workload imagepngdecode -AffinityMask 0xFFFF
./Build/Run-LibraryComparisonBenchmarks.ps1 -Profile full -Workload imagetifflzw -AffinityMask 0xFFFF
./Build/Run-LibraryComparisonBenchmarks.ps1 -Profile full -Workload imagewebplossless -AffinityMask 0xFFFF
```

The runner records the mask in benchmark provenance. Affinity reduces scheduler noise; it does not make small differences meaningful, so conclusions should still favor effect sizes that clearly exceed run-to-run variation.

Every timed decode returns the same managed RGBA output shape. BenchmarkDotNet's allocation column still reports managed allocations only. SkiaSharp and Magick.NET perform material work in native code, so their managed allocation numbers are not total-memory comparisons; use process/native-memory profiling before drawing memory-efficiency conclusions across engines.
