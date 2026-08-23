# OfficeIMO image benchmarks

This project measures the first-party `OfficeIMO.Core` image engine without adding image-library dependencies to the product. The corpus covers PNG, JPEG, GIF, TIFF, and BMP metadata and decode paths, deterministic RGBA encoding to PNG/JPEG/TIFF/WebP, bilinear resize, and placement-aware optimization. Encode and transform workloads run at several image or target sizes so managed-allocation growth remains visible instead of being inferred from one small sample.

The validation pass reports encoded bytes separately from benchmark time. It includes PNG compression, JPEG quality/subsampling/progressive modes with MAE and PSNR, TIFF compression, literal-lossless WebP, and every supported static source-to-output conversion. Lossless rows require exact RGBA equality. JPEG rows compare against the alpha-flattened source. Animated input is rejected by the optimization matrix rather than silently reduced to one frame.

The repository's `snail.bmp` has four trailing bytes beyond its declared BMP file size, so it remains a metadata fixture; decode measurements use a generated canonical 24-bit BMP instead of weakening the decoder's strict container contract.

Every timed workload is validated in global setup. Run the complete validation pass before collecting measurements:

```powershell
dotnet run --project OfficeIMO.Drawing.Benchmarks -c Release -f net10.0 -- --validate
```

Start benchmark work with a short diagnostic run:

```powershell
dotnet run --project OfficeIMO.Drawing.Benchmarks -c Release -f net10.0 -- --job Dry --filter '*ImageEncodeBenchmarks*'
```

Use a normal BenchmarkDotNet job only after the workload, output dimensions, format, and pixel-preservation contract have been validated. Benchmark artifacts belong in an ignored or temporary output directory and should not be committed.

For a heterogeneous or multi-CCD processor, pin the run to a reviewed processor region. For example, use this mask when logical processors 0-15 form one representative region:

```powershell
dotnet run --project OfficeIMO.Drawing.Benchmarks -c Release -f net10.0 -- --job Short --filter '*ImageEncodeBenchmarks*' --affinity 65535
```

Interpret the columns as separate tradeoffs:

- `Allocated` is managed allocation per operation. It does not include native allocations from comparison libraries.
- Encoded byte length and JPEG fidelity come from deterministic validation, not timed iterations.
- Small timing differences are not decisions on their own, even with affinity. Prefer changes that also remove whole image-sized buffers, reduce output materially, or improve a correctness contract.
- JPEG 4:2:0, optimized Huffman tables, progressive scans, TIFF Deflate, and TIFF PackBits have different CPU, fidelity, allocation, and file-size profiles. Keep them explicit when no one policy wins every axis.
- OfficeIMO WebP currently emits a standards-compatible literal-only lossless VP8L subset. It is fast and auditable, but its output size is close to raw RGBA; do not interpret it as a general-purpose WebP compression comparison.
