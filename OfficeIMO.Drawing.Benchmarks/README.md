# OfficeIMO image benchmarks

This project measures the first-party `OfficeIMO.Core` image engine without adding image-library dependencies to the product. The file corpus covers PNG, JPEG, GIF, TIFF, and BMP metadata and decode paths. Deterministic generated scenarios add tiny images, screenshots, text, line art, scans, alpha graphics, high-entropy pixels, a photo, and a 4096x3072 stress image. The suite covers RGBA encoding to PNG/JPEG/TIFF/WebP, caller-owned streaming output, bilinear, area, and Lanczos3 resize, and placement-aware optimization so allocation growth is not inferred from one logo or one synthetic pattern.

The validation pass reports encoded bytes separately from benchmark time. It includes PNG compression, JPEG quality/subsampling/progressive modes with MAE and PSNR, TIFF compression, literal-lossless WebP, and every supported static source-to-output conversion. Lossless rows require exact RGBA equality. JPEG rows compare against the alpha-flattened source. Animated input is rejected by the optimization matrix rather than silently reduced to one frame.

The resampling evidence uses exact pixel-area integration as the antialiasing reference for four-times downsampling. It reports premultiplied-RGB MAE, PSNR, alpha MAE, and a deterministic fingerprint for photo, text, line-art, and transparency fixtures. Premultiplied metrics avoid treating invisible RGB under zero alpha as visible error. These numbers describe visual tradeoffs; they are not interchangeable with elapsed time.

Write the same validated outputs for human visual inspection with an explicit artifact directory:

```powershell
dotnet run --project OfficeIMO.Drawing.Benchmarks -c Release -f net10.0 -- --resampling-previews Ignore/ImageResamplingPreviews
```

The repository's `snail.bmp` has four trailing bytes beyond its declared BMP file size, so it remains a metadata fixture; decode measurements use a generated canonical 24-bit BMP instead of weakening the decoder's strict container contract.

Every timed workload is validated in global setup. Run the complete validation pass before collecting measurements:

```powershell
dotnet run --project OfficeIMO.Drawing.Benchmarks -c Release -f net10.0 -- --validate
```

Collect isolated managed-allocation, peak working-set, and peak private-byte deltas for representative materialized-versus-streamed encodes with:

```powershell
dotnet run --project OfficeIMO.Drawing.Benchmarks -c Release -f net10.0 -- --memory-evidence
```

Pass one or more scenario names such as `Screenshot`, `HighEntropy`, or `VeryLarge` to narrow that run. Each row starts after the decoded/generated source image is resident. `Peak private` is a process-level managed-plus-native boundary, not a claim that the runtime can attribute every byte to a specific native codec.

Start benchmark work with a short diagnostic run:

```powershell
dotnet run --project OfficeIMO.Drawing.Benchmarks -c Release -f net10.0 -- --job Dry --filter '*ImageEncodeBenchmarks*'
```

Use `*ImageStreamingEncodeBenchmarks*` to compare the existing materialized `byte[]` contract with a caller-owned non-buffering stream across the representative timed corpus. The stream lane validates the same encoded format, dimensions, and decoded pixels in setup; it intentionally measures output ownership without allocating another full result array. It uses baseline JPEG settings so progressive and optimized-Huffman work does not obscure that comparison; those JPEG modes remain covered by the encoding size/fidelity validator.

Use `*ImageResamplingBenchmarks*` to compare bilinear, pixel-area, and Lanczos3 downsampling on the validated photo, text, line-art, and transparency fixtures. The modes perform different quality work, so interpret their time and allocation beside the resampling fidelity matrix rather than as a parity race.

Use a normal BenchmarkDotNet job only after the workload, output dimensions, format, and pixel-preservation contract have been validated. Benchmark artifacts belong in an ignored or temporary output directory and should not be committed.

For the 9950X3D2 benchmark host, pin the run to the reviewed `0xFFFF` processor region (decimal `65535`, logical processors 0-15):

```powershell
dotnet run --project OfficeIMO.Drawing.Benchmarks -c Release -f net10.0 -- --job Short --filter '*ImageEncodeBenchmarks*' --affinity 65535
```

Interpret the columns as separate tradeoffs:

- `Allocated` is managed allocation per operation. It does not include native allocations from comparison libraries.
- Encoded byte length and JPEG fidelity come from deterministic validation, not timed iterations.
- Small timing differences are not decisions on their own, even with affinity. Prefer changes that also remove whole image-sized buffers, reduce output materially, or improve a correctness contract.
- JPEG 4:2:0, optimized Huffman tables, progressive scans, TIFF Deflate, and TIFF PackBits have different CPU, fidelity, allocation, and file-size profiles. Keep them explicit when no one policy wins every axis.
- OfficeIMO WebP currently emits a standards-compatible literal-only lossless VP8L subset. It is fast and auditable, but its output size is close to raw RGBA; do not interpret it as a general-purpose WebP compression comparison.
