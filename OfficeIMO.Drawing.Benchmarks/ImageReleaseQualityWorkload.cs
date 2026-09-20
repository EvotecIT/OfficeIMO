using System.Diagnostics;
using System.Security.Cryptography;

namespace OfficeIMO.Drawing.Benchmarks;

/// <summary>Provenance-bound image workload consumed by the PowerForge release gate.</summary>
public sealed class ImageReleaseQualityWorkload {
    private const long MaximumEncodedBytes = 256L * 1024L * 1024L;
    private readonly OfficeRasterImage _source;
    private readonly OfficeImageExportFormat _format;
    private readonly ImageEvidenceOperation _operation;
    private readonly OfficeRasterEncodingOptions _options;
    private readonly byte[] _input;
    private object? _result;

    /// <summary>Creates a deterministic corpus, format, and operation tuple.</summary>
    public ImageReleaseQualityWorkload(string scenarioId, string format, string operation) {
        ImageBenchmarkScenario scenario = ImageBenchmarkScenarios.Get(scenarioId);
        if (!Enum.TryParse(format, true, out _format) || !_format.IsRaster())
            throw new ArgumentException("The format must be a supported raster format.", nameof(format));
        if (!Enum.TryParse(operation, true, out _operation))
            throw new ArgumentException("The image evidence operation is invalid.", nameof(operation));
        ScenarioId = scenario.Id;
        Format = _format.ToString();
        Operation = _operation.ToString();
        _source = scenario.CreateImage();
        _options = CreateOptions();
        _input = OfficeRasterImageEncoder.Encode(_source, _format, _options, MaximumEncodedBytes);
        InputSha256 = Convert.ToHexString(SHA256.HashData(_input));
    }

    /// <summary>Deterministic corpus scenario.</summary>
    public string ScenarioId { get; }
    /// <summary>Raster container format.</summary>
    public string Format { get; }
    /// <summary>Measured operation.</summary>
    public string Operation { get; }
    /// <summary>SHA-256 identity of the immutable encoded input.</summary>
    public string InputSha256 { get; }
    /// <summary>SHA-256 identity of the validated output.</summary>
    public string OutputSha256 { get; private set; } = string.Empty;
    /// <summary>Relevant encoded input or output length.</summary>
    public long EncodedBytes { get; private set; }
    /// <summary>Managed allocation during the measured operation.</summary>
    public long ManagedAllocatedBytes { get; private set; }
    /// <summary>Peak process working-set growth during the measured operation.</summary>
    public long PeakWorkingSetBytes { get; private set; }
    /// <summary>Peak process private-byte growth during the measured operation.</summary>
    public long PeakPrivateBytes { get; private set; }
    /// <summary>Private-byte growth remaining after observed managed-heap growth.</summary>
    public long PeakNativeBytesEstimate { get; private set; }
    /// <summary>One when a repeated operation produces the same logical output.</summary>
    public int Deterministic { get; private set; }
    /// <summary>Cancellation response in milliseconds, or -1 where cancellation is not exposed.</summary>
    public double CancellationLatencyMilliseconds { get; private set; } = -1D;
    /// <summary>Mean absolute RGB error for the validated output.</summary>
    public double MeanAbsoluteError { get; private set; }

    /// <summary>Runs only the selected operation and captures allocation and process-memory peaks.</summary>
    public void Execute() {
        using var sampler = new ImageProcessMemorySampler();
        sampler.Start();
        long allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
        _result = _operation switch {
            ImageEvidenceOperation.Encode => OfficeRasterImageEncoder.Encode(_source, _format, _options, MaximumEncodedBytes),
            ImageEvidenceOperation.Decode => Decode(_input),
            ImageEvidenceOperation.Metadata => OfficeImageReader.Identify(_input),
            ImageEvidenceOperation.Optimize => OfficeImageOptimizer.Optimize(_input, CreateOptimizationRequest(), ScenarioId + "." + Format),
            ImageEvidenceOperation.Resample => Resize(_source),
            _ => throw new InvalidOperationException("Unsupported image evidence operation.")
        };
        ManagedAllocatedBytes = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;
        sampler.Stop();
        PeakWorkingSetBytes = sampler.PeakWorkingSetDelta;
        PeakPrivateBytes = sampler.PeakPrivateBytesDelta;
        PeakNativeBytesEstimate = sampler.PeakNativeBytesEstimate;
    }

    /// <summary>Validates fidelity, determinism, bounds, output identity, and cancellation where exposed.</summary>
    public void Validate() {
        if (_result == null) throw new InvalidOperationException("Execute must run before validation.");
        switch (_operation) {
            case ImageEvidenceOperation.Encode: ValidateEncode((byte[])_result); break;
            case ImageEvidenceOperation.Decode: ValidateDecode((OfficeRasterImage)_result); break;
            case ImageEvidenceOperation.Metadata: ValidateMetadata((OfficeImageInfo)_result); break;
            case ImageEvidenceOperation.Optimize: ValidateOptimize((OfficeImageOptimizationResult)_result); break;
            case ImageEvidenceOperation.Resample: ValidateResize((OfficeRasterImage)_result); break;
        }
        if (Deterministic != 1) throw new InvalidOperationException($"{ScenarioId} {Format} {Operation} was not deterministic.");
    }

    private void ValidateEncode(byte[] encoded) {
        ValidateFidelity(_source, Decode(encoded));
        byte[] repeated = OfficeRasterImageEncoder.Encode(_source, _format, _options, MaximumEncodedBytes);
        RecordEncoded(encoded, encoded.AsSpan().SequenceEqual(repeated));
        OfficeImageExportBatchLimitException limit = ExpectException<OfficeImageExportBatchLimitException>(() =>
            OfficeRasterImageEncoder.Encode(_source, _format, _options, 8L));
        if (limit.LimitName != nameof(OfficeImageExportOptions.MaximumTotalEncodedBytes))
            throw new InvalidOperationException("The bounded encoder reported the wrong public limit.");
        CancellationLatencyMilliseconds = MeasureEncodeCancellation();
    }

    private void ValidateDecode(OfficeRasterImage decoded) {
        OfficeRasterImage repeated = Decode(_input);
        OfficeRasterImage expected = _format == OfficeImageExportFormat.Jpeg ? repeated : _source;
        ValidateFidelity(expected, decoded);
        EncodedBytes = _input.LongLength;
        OutputSha256 = PixelHash(decoded);
        Deterministic = decoded.GetPixels().AsSpan().SequenceEqual(repeated.GetPixels()) ? 1 : 0;
        if (_format is OfficeImageExportFormat.Tiff or OfficeImageExportFormat.Webp)
            CancellationLatencyMilliseconds = ImageCancellationEvidence.Measure(_format).TotalMilliseconds;
    }

    private void ValidateMetadata(OfficeImageInfo info) {
        if (info.Format != ToImageFormat(_format) || info.Width != _source.Width || info.Height != _source.Height)
            throw new InvalidOperationException($"{ScenarioId} {Format} metadata did not match the encoded input.");
        EncodedBytes = _input.LongLength;
        OutputSha256 = MetadataHash(info);
        Deterministic = OutputSha256 == MetadataHash(OfficeImageReader.Identify(_input)) ? 1 : 0;
    }

    private void ValidateOptimize(OfficeImageOptimizationResult result) {
        if (result.Status != OfficeImageOptimizationStatus.Optimized || result.Metadata.HasLoss)
            throw new InvalidOperationException($"{ScenarioId} {Format} optimization returned {result.Status} or lost requested metadata.");
        byte[] output = result.Bytes;
        ValidateFidelity(Resize(Decode(_input)), Decode(output));
        byte[] repeated = OfficeImageOptimizer.Optimize(_input, CreateOptimizationRequest(), ScenarioId + "." + Format).Bytes;
        RecordEncoded(output, output.AsSpan().SequenceEqual(repeated));
    }

    private void ValidateResize(OfficeRasterImage resized) {
        byte[] pixels = resized.GetPixels();
        byte[] repeated = Resize(_source).GetPixels();
        EncodedBytes = pixels.LongLength;
        OutputSha256 = Convert.ToHexString(SHA256.HashData(pixels));
        Deterministic = pixels.AsSpan().SequenceEqual(repeated) ? 1 : 0;
    }

    private void RecordEncoded(byte[] bytes, bool deterministic) {
        EncodedBytes = bytes.LongLength;
        OutputSha256 = Convert.ToHexString(SHA256.HashData(bytes));
        Deterministic = deterministic ? 1 : 0;
    }

    private void ValidateFidelity(OfficeRasterImage expected, OfficeRasterImage actual) {
        if (expected.Width != actual.Width || expected.Height != actual.Height)
            throw new InvalidOperationException($"{ScenarioId} {Format} produced unexpected dimensions.");
        MeanAbsoluteError = CalculateMeanAbsoluteRgbError(expected, actual);
        if (_format is OfficeImageExportFormat.Png or OfficeImageExportFormat.Tiff or OfficeImageExportFormat.Webp) {
            if (!expected.GetPixels().AsSpan().SequenceEqual(actual.GetPixels()))
                throw new InvalidOperationException($"{ScenarioId} {Format} did not preserve lossless pixels.");
        } else if (MeanAbsoluteError > 18D) {
            throw new InvalidOperationException($"{ScenarioId} JPEG mean absolute RGB error {MeanAbsoluteError:F3} exceeded 18.");
        }
    }

    private double MeasureEncodeCancellation() {
        OfficeRasterEncodingCheckpoint expected = _format switch {
            OfficeImageExportFormat.Png => OfficeRasterEncodingCheckpoint.PngCompressionRow,
            OfficeImageExportFormat.Jpeg => OfficeRasterEncodingCheckpoint.JpegCoefficientRow,
            OfficeImageExportFormat.Tiff => OfficeRasterEncodingCheckpoint.TiffCompressionRow,
            OfficeImageExportFormat.Webp => OfficeRasterEncodingCheckpoint.WebpCompressionBlock,
            _ => throw new ArgumentOutOfRangeException(nameof(_format))
        };
        using var started = new ManualResetEventSlim();
        using var requested = new ManualResetEventSlim();
        using var cancellation = new CancellationTokenSource();
        long requestedAt = 0L;
        var thread = new Thread(() => { started.Wait(); Volatile.Write(ref requestedAt, Stopwatch.GetTimestamp()); cancellation.Cancel(); requested.Set(); }) { IsBackground = true };
        thread.Start();
        int checkpoints = 0;
        try {
            using var output = new MemoryStream();
            ExpectException<OperationCanceledException>(() => OfficeRasterImageEncoder.EncodeTo(
                _source, _format, output, _options, long.MaxValue, cancellation.Token,
                checkpoint => {
                    if (checkpoint != expected || Interlocked.Increment(ref checkpoints) != 2) return;
                    started.Set();
                    if (!requested.Wait(TimeSpan.FromSeconds(5))) throw new InvalidOperationException("Cancellation request did not arrive.");
                }));
        } finally { started.Set(); thread.Join(); }
        if (checkpoints < 2 || requestedAt == 0L) throw new InvalidOperationException("The encoder did not reach its cancellation checkpoint.");
        return Stopwatch.GetElapsedTime(requestedAt).TotalMilliseconds;
    }

    private OfficeImageOptimizationRequest CreateOptimizationRequest() => new(Math.Max(1, _source.Width / 2), Math.Max(1, _source.Height / 2)) {
        PreserveAspectRatio = false,
        ResamplingMode = OfficeRasterResamplingMode.Lanczos3,
        OutputFormat = ToImageFormat(_format),
        KeepOriginalWhenNotSmaller = false,
        MetadataPolicy = _format == OfficeImageExportFormat.Webp
            ? OfficeImageMetadataPolicy.SelectiveCopy
            : OfficeImageMetadataPolicy.Preserve,
        MetadataSelection = _format == OfficeImageExportFormat.Webp
            ? OfficeImageMetadataKinds.None
            : OfficeImageMetadataKinds.All,
        JpegQuality = 85,
        JpegSubsampling = OfficeJpegSubsampling.Y420,
        TiffCompression = OfficeTiffCompression.PackBits
    };

    private static OfficeRasterImage Resize(OfficeRasterImage source) => OfficeRasterResampler.Resize(
        source, Math.Max(1, source.Width / 2), Math.Max(1, source.Height / 2), OfficeRasterResamplingMode.Lanczos3);

    private static OfficeRasterImage Decode(byte[] bytes) {
        if (!OfficeRasterImageDecoder.TryDecode(bytes, out OfficeRasterImage? image) || image == null)
            throw new InvalidOperationException("The release-quality corpus input could not be decoded.");
        return image;
    }

    private static OfficeImageFormat ToImageFormat(OfficeImageExportFormat format) => format switch {
        OfficeImageExportFormat.Png => OfficeImageFormat.Png,
        OfficeImageExportFormat.Jpeg => OfficeImageFormat.Jpeg,
        OfficeImageExportFormat.Tiff => OfficeImageFormat.Tiff,
        OfficeImageExportFormat.Webp => OfficeImageFormat.Webp,
        _ => throw new ArgumentOutOfRangeException(nameof(format))
    };

    private static OfficeRasterEncodingOptions CreateOptions() => new() {
        DpiX = 144D, DpiY = 120D,
        Png = new OfficePngEncodeOptions { Compression = OfficePngCompression.Optimal },
        Jpeg = new OfficeJpegEncodeOptions { Quality = 85, Subsampling = OfficeJpegSubsampling.Y420, Background = OfficeColor.White },
        Tiff = new OfficeTiffEncodeOptions { Compression = OfficeTiffCompression.PackBits }
    };

    private static string PixelHash(OfficeRasterImage image) => Convert.ToHexString(SHA256.HashData(image.GetPixels()));
    private static string MetadataHash(OfficeImageInfo info) => Convert.ToHexString(SHA256.HashData(
        System.Text.Encoding.UTF8.GetBytes($"{info.Format}|{info.Width}|{info.Height}|{info.DpiX:R}|{info.DpiY:R}")));
    private static TException ExpectException<TException>(Action action) where TException : Exception {
        try { action(); } catch (TException exception) { return exception; }
        throw new InvalidOperationException($"Expected {typeof(TException).Name}.");
    }
    private static double CalculateMeanAbsoluteRgbError(OfficeRasterImage expected, OfficeRasterImage actual) {
        byte[] left = expected.GetPixels(); byte[] right = actual.GetPixels(); long total = 0L;
        for (int i = 0; i < left.Length; i += 4) {
            total += Math.Abs(left[i] - right[i]); total += Math.Abs(left[i + 1] - right[i + 1]); total += Math.Abs(left[i + 2] - right[i + 2]);
        }
        return left.Length == 0 ? 0D : total / (left.Length / 4D * 3D);
    }

    private enum ImageEvidenceOperation { Encode, Decode, Metadata, Optimize, Resample }
}
