using System.Diagnostics;
using System.Buffers.Binary;
using System.Globalization;
using System.Security.Cryptography;
using System.Text;

namespace OfficeIMO.Drawing.Benchmarks;

/// <summary>Provenance-bound image workload consumed by the PowerForge release gate.</summary>
public sealed class ImageReleaseQualityWorkload {
    private const long MaximumEncodedBytes = 256L * 1024L * 1024L;
    private readonly OfficeRasterImage _source;
    private readonly OfficeImageExportFormat _format;
    private readonly ImageEvidenceOperation _operation;
    private readonly OfficeRasterEncodingOptions _options;
    private readonly byte[] _input;
    private ImageProcessMemorySampler? _sampler;
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
        byte[] inputHash = SHA256.HashData(_input);
        InputSha256 = Convert.ToHexString(inputHash);
        InputHashWord0 = BinaryPrimitives.ReadUInt32BigEndian(inputHash.AsSpan(0, 4));
        InputHashWord1 = BinaryPrimitives.ReadUInt32BigEndian(inputHash.AsSpan(4, 4));
        InputHashWord2 = BinaryPrimitives.ReadUInt32BigEndian(inputHash.AsSpan(8, 4));
        InputHashWord3 = BinaryPrimitives.ReadUInt32BigEndian(inputHash.AsSpan(12, 4));
        InputHashWord4 = BinaryPrimitives.ReadUInt32BigEndian(inputHash.AsSpan(16, 4));
        InputHashWord5 = BinaryPrimitives.ReadUInt32BigEndian(inputHash.AsSpan(20, 4));
        InputHashWord6 = BinaryPrimitives.ReadUInt32BigEndian(inputHash.AsSpan(24, 4));
        InputHashWord7 = BinaryPrimitives.ReadUInt32BigEndian(inputHash.AsSpan(28, 4));
        byte[] provenanceHash = CreateWorkloadProvenanceHash();
        ProvenanceSha256 = Convert.ToHexString(provenanceHash);
        ProvenanceHashWord0 = BinaryPrimitives.ReadUInt32BigEndian(provenanceHash.AsSpan(0, 4));
        ProvenanceHashWord1 = BinaryPrimitives.ReadUInt32BigEndian(provenanceHash.AsSpan(4, 4));
        ProvenanceHashWord2 = BinaryPrimitives.ReadUInt32BigEndian(provenanceHash.AsSpan(8, 4));
        ProvenanceHashWord3 = BinaryPrimitives.ReadUInt32BigEndian(provenanceHash.AsSpan(12, 4));
        ProvenanceHashWord4 = BinaryPrimitives.ReadUInt32BigEndian(provenanceHash.AsSpan(16, 4));
        ProvenanceHashWord5 = BinaryPrimitives.ReadUInt32BigEndian(provenanceHash.AsSpan(20, 4));
        ProvenanceHashWord6 = BinaryPrimitives.ReadUInt32BigEndian(provenanceHash.AsSpan(24, 4));
        ProvenanceHashWord7 = BinaryPrimitives.ReadUInt32BigEndian(provenanceHash.AsSpan(28, 4));
    }

    /// <summary>Deterministic corpus scenario.</summary>
    public string ScenarioId { get; }
    /// <summary>Raster container format.</summary>
    public string Format { get; }
    /// <summary>Measured operation.</summary>
    public string Operation { get; }
    /// <summary>SHA-256 identity of the immutable encoded input.</summary>
    public string InputSha256 { get; }
    /// <summary>First 32-bit word of the immutable input hash, for numeric benchmark evidence.</summary>
    public long InputHashWord0 { get; }
    /// <summary>Second 32-bit word of the immutable input hash, for numeric benchmark evidence.</summary>
    public long InputHashWord1 { get; }
    /// <summary>Third 32-bit word of the immutable input hash, for numeric benchmark evidence.</summary>
    public long InputHashWord2 { get; }
    /// <summary>Fourth 32-bit word of the immutable input hash, for numeric benchmark evidence.</summary>
    public long InputHashWord3 { get; }
    /// <summary>Fifth 32-bit word of the immutable input hash, for numeric benchmark evidence.</summary>
    public long InputHashWord4 { get; }
    /// <summary>Sixth 32-bit word of the immutable input hash, for numeric benchmark evidence.</summary>
    public long InputHashWord5 { get; }
    /// <summary>Seventh 32-bit word of the immutable input hash, for numeric benchmark evidence.</summary>
    public long InputHashWord6 { get; }
    /// <summary>Eighth 32-bit word of the immutable input hash, for numeric benchmark evidence.</summary>
    public long InputHashWord7 { get; }
    /// <summary>SHA-256 identity of the source pixels and workload configuration.</summary>
    public string ProvenanceSha256 { get; }
    /// <summary>First 32-bit word of the workload provenance hash.</summary>
    public long ProvenanceHashWord0 { get; }
    /// <summary>Second 32-bit word of the workload provenance hash.</summary>
    public long ProvenanceHashWord1 { get; }
    /// <summary>Third 32-bit word of the workload provenance hash.</summary>
    public long ProvenanceHashWord2 { get; }
    /// <summary>Fourth 32-bit word of the workload provenance hash.</summary>
    public long ProvenanceHashWord3 { get; }
    /// <summary>Fifth 32-bit word of the workload provenance hash.</summary>
    public long ProvenanceHashWord4 { get; }
    /// <summary>Sixth 32-bit word of the workload provenance hash.</summary>
    public long ProvenanceHashWord5 { get; }
    /// <summary>Seventh 32-bit word of the workload provenance hash.</summary>
    public long ProvenanceHashWord6 { get; }
    /// <summary>Eighth 32-bit word of the workload provenance hash.</summary>
    public long ProvenanceHashWord7 { get; }
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

    /// <summary>Starts process-memory sampling outside measured operation time.</summary>
    public void BeginMeasurement() {
        if (_sampler != null) throw new InvalidOperationException("Image evidence measurement is already active.");
        var sampler = new ImageProcessMemorySampler();
        try {
            sampler.Start();
            _sampler = sampler;
        } catch {
            sampler.Dispose();
            throw;
        }
    }

    /// <summary>Runs only the selected image operation.</summary>
    public void Execute() {
        long allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
        try {
            _result = _operation switch {
                ImageEvidenceOperation.Encode => OfficeRasterImageEncoder.Encode(_source, _format, _options, MaximumEncodedBytes),
                ImageEvidenceOperation.Decode => Decode(_input),
                ImageEvidenceOperation.Metadata => OfficeImageReader.Identify(_input),
                ImageEvidenceOperation.Optimize => OfficeImageOptimizer.Optimize(_input, CreateOptimizationRequest(), ScenarioId + "." + Format),
                ImageEvidenceOperation.Resample => Resize(_source),
                _ => throw new InvalidOperationException("Unsupported image evidence operation.")
            };
        } catch {
            ManagedAllocatedBytes = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;
            CompleteMeasurement();
            throw;
        }
        ManagedAllocatedBytes = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;
    }

    /// <summary>Stops sampling and publishes memory metrics outside measured operation time.</summary>
    public void CompleteMeasurement() {
        ImageProcessMemorySampler? sampler = Interlocked.Exchange(ref _sampler, null);
        if (sampler == null) return;
        try {
            sampler.Stop();
            PeakWorkingSetBytes = sampler.PeakWorkingSetDelta;
            PeakPrivateBytes = sampler.PeakPrivateBytesDelta;
            PeakNativeBytesEstimate = sampler.PeakNativeBytesEstimate;
        } finally {
            sampler.Dispose();
        }
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
        ValidateFidelity(_source, decoded);
        EncodedBytes = _input.LongLength;
        OutputSha256 = PixelHash(decoded);
        Deterministic = decoded.GetPixels().AsSpan().SequenceEqual(repeated.GetPixels()) ? 1 : 0;
        if (_format is OfficeImageExportFormat.Tiff or OfficeImageExportFormat.Webp)
            CancellationLatencyMilliseconds = ImageCancellationEvidence.Measure(_format).TotalMilliseconds;
    }

    private void ValidateMetadata(OfficeImageInfo info) {
        (double expectedDpiX, double expectedDpiY) = GetExpectedEncodedDpi(_format);
        if (info.Format != ToImageFormat(_format) || info.Width != _source.Width || info.Height != _source.Height ||
            Math.Abs(info.DpiX - expectedDpiX) > 0.0001D || Math.Abs(info.DpiY - expectedDpiY) > 0.0001D)
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
        int expectedWidth = Math.Max(1, _source.Width / 2);
        int expectedHeight = Math.Max(1, _source.Height / 2);
        if (resized.Width != expectedWidth || resized.Height != expectedHeight)
            throw new InvalidOperationException($"{ScenarioId} resampling produced {resized.Width}x{resized.Height}; expected {expectedWidth}x{expectedHeight}.");

        byte[] pixels = resized.GetPixels();
        byte[] repeated = Resize(_source).GetPixels();
        (double sourceRed, double sourceGreen, double sourceBlue, double sourceAlpha) = CalculateChannelMeans(_source.GetPixels());
        (double resultRed, double resultGreen, double resultBlue, double resultAlpha) = CalculateChannelMeans(pixels);
        MeanAbsoluteError = (Math.Abs(sourceRed - resultRed) + Math.Abs(sourceGreen - resultGreen) + Math.Abs(sourceBlue - resultBlue)) / 3D;
        if (MeanAbsoluteError > 8D)
            throw new InvalidOperationException($"{ScenarioId} resampling mean-channel RGB drift {MeanAbsoluteError:F3} exceeded 8.");
        if (Math.Abs(sourceAlpha - resultAlpha) > 8D)
            throw new InvalidOperationException($"{ScenarioId} resampling alpha-coverage drift exceeded 8.");

        ValidateDynamicRange(pixels);
        if (HasTransparentAndOpaqueSamples(_source.GetPixels()) && !HasTransparentAndOpaqueSamples(pixels))
            throw new InvalidOperationException($"{ScenarioId} resampling did not preserve both transparent and opaque coverage.");

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

    private static (double DpiX, double DpiY) GetExpectedEncodedDpi(OfficeImageExportFormat format) {
        // PNG pHYs stores whole pixels per metre, so the fixed 144x120 corpus density is quantized.
        return format switch {
            OfficeImageExportFormat.Png => (5669D * 0.0254D, 4724D * 0.0254D),
            OfficeImageExportFormat.Jpeg or OfficeImageExportFormat.Tiff or OfficeImageExportFormat.Webp => (144D, 120D),
            _ => throw new ArgumentOutOfRangeException(nameof(format))
        };
    }

    private static OfficeRasterEncodingOptions CreateOptions() => new() {
        DpiX = 144D, DpiY = 120D,
        Png = new OfficePngEncodeOptions { Compression = OfficePngCompression.Optimal },
        Jpeg = new OfficeJpegEncodeOptions { Quality = 85, Subsampling = OfficeJpegSubsampling.Y420, Background = OfficeColor.White },
        Tiff = new OfficeTiffEncodeOptions { Compression = OfficeTiffCompression.PackBits }
    };

    private byte[] CreateWorkloadProvenanceHash() {
        OfficeColor background = _options.Jpeg.Background;
        string configuration = string.Join("|", new[] {
            "officeimo-image-evidence-v1",
            ScenarioId,
            _source.Width.ToString(CultureInfo.InvariantCulture),
            _source.Height.ToString(CultureInfo.InvariantCulture),
            Format,
            Operation,
            _options.WriteResolutionMetadata.ToString(),
            _options.DpiX.ToString("R", CultureInfo.InvariantCulture),
            _options.DpiY.ToString("R", CultureInfo.InvariantCulture),
            _options.Png.Compression.ToString(),
            _options.Jpeg.Quality.ToString(CultureInfo.InvariantCulture),
            _options.Jpeg.Subsampling.ToString(),
            _options.Jpeg.Progressive.ToString(),
            _options.Jpeg.OptimizeHuffman.ToString(),
            $"{background.R},{background.G},{background.B},{background.A}",
            _options.Tiff.Compression.ToString(),
            _options.Tiff.Predictor.ToString(),
            OfficeRasterResamplingMode.Lanczos3.ToString(),
            MaximumEncodedBytes.ToString(CultureInfo.InvariantCulture),
            "optimize-half-size|preserve-aspect=false|keep-original=false|jpeg-quality=85|jpeg-subsampling=Y420|tiff-compression=PackBits"
        });
        using IncrementalHash hash = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        hash.AppendData(Encoding.UTF8.GetBytes(configuration));
        hash.AppendData(_source.GetPixels());
        return hash.GetHashAndReset();
    }

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

    private static (double Red, double Green, double Blue, double Alpha) CalculateChannelMeans(byte[] pixels) {
        if (pixels.Length == 0) return default;
        long red = 0L; long green = 0L; long blue = 0L; long alpha = 0L;
        for (int i = 0; i < pixels.Length; i += 4) {
            red += pixels[i]; green += pixels[i + 1]; blue += pixels[i + 2]; alpha += pixels[i + 3];
        }
        double count = pixels.Length / 4D;
        return (red / count, green / count, blue / count, alpha / count);
    }

    private static void ValidateDynamicRange(byte[] pixels) {
        byte minimum = byte.MaxValue;
        byte maximum = byte.MinValue;
        for (int i = 0; i < pixels.Length; i += 4) {
            minimum = Math.Min(minimum, Math.Min(pixels[i], Math.Min(pixels[i + 1], pixels[i + 2])));
            maximum = Math.Max(maximum, Math.Max(pixels[i], Math.Max(pixels[i + 1], pixels[i + 2])));
        }
        if (maximum - minimum < 8)
            throw new InvalidOperationException("Resampling collapsed the source image's visible RGB dynamic range.");
    }

    private static bool HasTransparentAndOpaqueSamples(byte[] pixels) {
        bool transparent = false;
        bool opaque = false;
        for (int i = 3; i < pixels.Length; i += 4) {
            transparent |= pixels[i] < 32;
            opaque |= pixels[i] > 223;
            if (transparent && opaque) return true;
        }
        return false;
    }

    private enum ImageEvidenceOperation { Encode, Decode, Metadata, Optimize, Resample }
}
