# OfficeIMO image benchmarks

This project measures the first-party `OfficeIMO.Core` image engine without adding image-library dependencies to the product. The corpus covers PNG, JPEG, GIF, TIFF, and BMP metadata and decode paths, deterministic RGBA encoding to PNG/JPEG/TIFF/WebP, bilinear resize, and placement-aware optimization. The repository's `snail.bmp` has four trailing bytes beyond its declared BMP file size, so it remains a metadata fixture; decode measurements use a generated canonical 24-bit BMP instead of weakening the decoder's strict container contract.

Every timed workload is validated in global setup. Run the complete validation pass before collecting measurements:

```powershell
dotnet run --project OfficeIMO.Drawing.Benchmarks -c Release -f net10.0 -- --validate
```

Start benchmark work with a short diagnostic run:

```powershell
dotnet run --project OfficeIMO.Drawing.Benchmarks -c Release -f net10.0 -- --job Dry --filter '*ImageEncodeBenchmarks*'
```

Use a normal BenchmarkDotNet job only after the workload, output dimensions, format, and pixel-preservation contract have been validated. Benchmark artifacts belong in an ignored or temporary output directory and should not be committed.
