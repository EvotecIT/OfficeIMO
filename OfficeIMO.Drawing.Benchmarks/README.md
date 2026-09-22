# OfficeIMO image benchmarks

This project measures the first-party `OfficeIMO.Core` image engine without adding image-library dependencies to the product. The file corpus covers PNG, JPEG, GIF, TIFF, and BMP metadata and decode paths. Deterministic generated scenarios add tiny images, screenshots, text, line art, scans, alpha graphics, high-entropy pixels, a photo, and a 4096x3072 stress image. The suite covers RGBA encoding to PNG/JPEG/TIFF/WebP, caller-owned streaming output, bilinear, area, and Lanczos3 resize, and placement-aware optimization so allocation growth is not inferred from one logo or one synthetic pattern.

The validation pass reports encoded bytes separately from benchmark time. It includes PNG compression, JPEG quality/subsampling/progressive modes with MAE and PSNR, TIFF LZW/PackBits/Deflate, lossless VP8L WebP, and every supported static source-to-output conversion. Lossless rows require exact RGBA equality. JPEG rows compare against the alpha-flattened source. Animated and multi-page input is rejected by the optimizer rather than silently reduced to one frame or page.

The same pass schedules cancellation during bounded TIFF and WebP decode of a 4096x1025 image, records the observed latency, and rejects a run that does not observe cancellation within two seconds. This is a regression ceiling rather than a throughput benchmark; compare repeatable latency changes separately from encoding and fidelity results.

The resampling evidence uses exact pixel-area integration as the antialiasing reference for four-times downsampling. It reports premultiplied-RGB MAE, PSNR, alpha MAE, and a deterministic fingerprint for photo, text, line-art, and transparency fixtures. Encoded-sRGB and optional linear-light Area/Lanczos results are evaluated against the matching color-space reference. Premultiplied metrics avoid treating invisible RGB under zero alpha as visible error. These numbers describe visual tradeoffs; they are not interchangeable with elapsed time.

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

## Release-quality evidence suite

The repository-level release gate uses the same benchmark assembly through PowerForge's structured evidence runner. Build the benchmark project first, then run the suite from the repository root with PowerShell on .NET 10 or newer and PSPublishModule 3.0.141 or later:

```powershell
dotnet build OfficeIMO.Drawing.Benchmarks/OfficeIMO.Drawing.Benchmarks.csproj -c Release -f net10.0
pwsh ./Build/Benchmarks/Run-ReleaseQualityImageEvidence.ps1
```

Use `-Plan` to inspect the resolved cases and policy without executing measurements. `-BinaryRoot`, `-OutputRoot`, `-WarmupCount`, and `-IterationCount` provide explicit inputs for a controlled run. The default output is `.validation/release-quality-images`; it contains JSON, CSV, and Markdown evidence and remains a task-owned validation artifact rather than a committed benchmark result.

The suite runs representative PNG, JPEG, TIFF, and WebP inputs through encode, decode, metadata,
and placement-optimization workloads, plus generated line-art, text, and transparency through
Lanczos resampling. Each case records the encoded-input hash, a source-pixel and workload-configuration
provenance hash, and the benchmark-assembly hash, then validates dimensions, fidelity, deterministic
output, and bounded cancellation where that API exposes cancellation. Regression comparisons require
matching workload provenance for every operation, including encode and resample.
Evidence separates elapsed time, encoded bytes, managed allocation, peak working set, peak private
bytes, a diagnostic private-minus-live-managed-heap estimate, cancellation latency, and mean absolute
error. Memory sampling starts before and stops after the PowerForge-timed operation, so sampler thread
startup and shutdown are excluded from elapsed time. Managed allocation is captured immediately
around the workload operation on its executing thread, excluding host dispatch and process-sampler
setup or teardown. A non-planning run fails when any case is not
`Succeeded`.

Pull-request runs download the matching operating-system artifact from the latest successful
`master` workflow and use PowerForge `Test-BenchmarkGate` comparisons. The first run permits new
scenario keys; after they are present on `master`, subsequent Windows, Linux, and macOS runs gate
the full matrix. Process-memory gates use an absolute noise allowance as well as relative
tolerance because a zero baseline is common for short operations. Hosted CI gates validated output,
managed allocation, peak working set, peak private bytes, cancellation, and image error. It reports
elapsed time and the native estimate without gating them: hosted-runner timing varies between runs,
and subtracting the live managed heap from process private bytes cannot reliably attribute memory
to native allocations when garbage collection changes the live heap. Use controlled repeat runs
to investigate those diagnostics. Pass a reference summary when
reproducing the same comparison locally:

```powershell
pwsh ./Build/Benchmarks/Run-ReleaseQualityImageEvidence.ps1 `
    -ReferenceSummaryPath ./reference/summary.json
```

TIFF or WebP decoder optimization is justified only by a large, repeatable decode gap in the full
corpus on every supported operating system. A single-machine quick run is diagnostic evidence, not
an optimization priority.

Interpret correctness and cancellation as gates. Compare encoded size, elapsed time, allocation, and peak memory as separate tradeoffs across equivalent cases and the same machine/runtime; do not treat the fastest row alone as release-quality evidence. JPEG error is a lossy-fidelity metric, while the lossless cases must satisfy their exact decoded-pixel contract.

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
- OfficeIMO WebP deterministically selects a literal or bounded prediction/subtract-green/LZ77/Huffman VP8L stream. Interpret compression size beside equivalent external validation; lossy and animated WebP are not part of this lane.
