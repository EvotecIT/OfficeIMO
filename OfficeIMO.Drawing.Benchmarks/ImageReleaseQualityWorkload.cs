using System.Security.Cryptography;

namespace OfficeIMO.Drawing.Benchmarks;

/// <summary>
/// Reproducible bounded-encoder workload consumed by PowerForge release-quality evidence.
/// </summary>
public sealed class ImageReleaseQualityWorkload {
    private readonly OfficeRasterImage _source;
    private readonly OfficeImageExportFormat _format;
    private readonly OfficeRasterEncodingOptions _options;
    private byte[]? _encoded;

    /// <summary>Creates one deterministic scenario and encoder format pair.</summary>
    public ImageReleaseQualityWorkload(string scenarioId, string format) {
        ImageBenchmarkScenario scenario = ImageBenchmarkScenarios.Get(scenarioId);
        if (!Enum.TryParse(format, ignoreCase: true, out _format) || !_format.IsRaster()) {
            throw new ArgumentException("The release-quality format must be a raster format.", nameof(format));
        }

        ScenarioId = scenario.Id;
        Format = _format.ToString();
        _source = scenario.CreateImage();
        _options = new OfficeRasterEncodingOptions {
            DpiX = 144D,
            DpiY = 120D,
            Png = new OfficePngEncodeOptions { Compression = OfficePngCompression.Optimal },
            Jpeg = new OfficeJpegEncodeOptions {
                Quality = 85,
                Subsampling = OfficeJpegSubsampling.Y420,
                Progressive = false,
                OptimizeHuffman = false,
                Background = OfficeColor.White
            },
            Tiff = new OfficeTiffEncodeOptions { Compression = OfficeTiffCompression.PackBits }
        };
    }

    /// <summary>Deterministic scenario identifier.</summary>
    public string ScenarioId { get; }

    /// <summary>Raster format name.</summary>
    public string Format { get; }

    /// <summary>Length of the last validated encoded result.</summary>
    public long EncodedBytes => _encoded?.LongLength ?? 0L;

    /// <summary>Managed bytes allocated by the last bounded encode call.</summary>
    public long ManagedAllocatedBytes { get; private set; }

    /// <summary>SHA-256 of the last validated encoded result.</summary>
    public string EncodedSha256 { get; private set; } = string.Empty;

    /// <summary>One when repeated encoding produced byte-identical output.</summary>
    public int Deterministic { get; private set; }

    /// <summary>One when the public bounded encoder observed cancellation.</summary>
    public int CancellationObserved { get; private set; }

    /// <summary>Mean absolute RGB error for lossy output; zero for exact lossless output.</summary>
    public double MeanAbsoluteError { get; private set; }

    /// <summary>Runs the bounded materialized encoder.</summary>
    public void Encode() {
        long allocatedBefore = GC.GetAllocatedBytesForCurrentThread();
        _encoded = OfficeRasterImageEncoder.Encode(
            _source,
            _format,
            _options,
            maximumEncodedBytes: 256L * 1024L * 1024L,
            cancellationToken: CancellationToken.None);
        ManagedAllocatedBytes = GC.GetAllocatedBytesForCurrentThread() - allocatedBefore;
    }

    /// <summary>Validates dimensions, fidelity, determinism, size limits, and cancellation.</summary>
    public void Validate() {
        byte[] encoded = _encoded ?? throw new InvalidOperationException("Encode must run before validation.");
        if (!OfficeRasterImageDecoder.TryDecode(encoded, out OfficeRasterImage? decoded) || decoded == null) {
            throw new InvalidOperationException($"{ScenarioId} {Format} could not be decoded.");
        }
        if (decoded.Width != _source.Width || decoded.Height != _source.Height) {
            throw new InvalidOperationException(
                $"{ScenarioId} {Format} decoded to {decoded.Width}x{decoded.Height}; expected {_source.Width}x{_source.Height}.");
        }

        MeanAbsoluteError = CalculateMeanAbsoluteRgbError(_source, decoded);
        if (_format is OfficeImageExportFormat.Png or OfficeImageExportFormat.Tiff or OfficeImageExportFormat.Webp) {
            if (!_source.GetPixels().AsSpan().SequenceEqual(decoded.GetPixels())) {
                throw new InvalidOperationException($"{ScenarioId} {Format} did not preserve lossless pixels.");
            }
        } else if (MeanAbsoluteError > 18D) {
            throw new InvalidOperationException(
                $"{ScenarioId} {Format} mean absolute RGB error {MeanAbsoluteError:F3} exceeded 18.");
        }

        byte[] repeated = OfficeRasterImageEncoder.Encode(
            _source,
            _format,
            _options,
            maximumEncodedBytes: 256L * 1024L * 1024L,
            cancellationToken: CancellationToken.None);
        Deterministic = encoded.AsSpan().SequenceEqual(repeated) ? 1 : 0;
        if (Deterministic != 1) {
            throw new InvalidOperationException($"{ScenarioId} {Format} output was not deterministic.");
        }

        EncodedSha256 = Convert.ToHexString(SHA256.HashData(encoded));
        OfficeImageExportBatchLimitException limit = ExpectException<OfficeImageExportBatchLimitException>(() =>
            OfficeRasterImageEncoder.Encode(
                _source,
                _format,
                _options,
                maximumEncodedBytes: 8L,
                cancellationToken: CancellationToken.None));
        if (limit.LimitName != nameof(OfficeImageExportOptions.MaximumTotalEncodedBytes)) {
            throw new InvalidOperationException("The bounded encoder reported the wrong public limit.");
        }

        ValidateInFlightCancellation();
        CancellationObserved = 1;
    }

    private void ValidateInFlightCancellation() {
        OfficeRasterEncodingCheckpoint expectedCheckpoint = _format switch {
            OfficeImageExportFormat.Png => OfficeRasterEncodingCheckpoint.PngCompressionRow,
            OfficeImageExportFormat.Jpeg => OfficeRasterEncodingCheckpoint.JpegCoefficientRow,
            OfficeImageExportFormat.Tiff => OfficeRasterEncodingCheckpoint.TiffCompressionRow,
            OfficeImageExportFormat.Webp => OfficeRasterEncodingCheckpoint.WebpCompressionBlock,
            _ => throw new ArgumentOutOfRangeException(nameof(_format))
        };
        using var encodingStarted = new ManualResetEventSlim();
        using var cancellationRequested = new ManualResetEventSlim();
        using var cancellation = new CancellationTokenSource();
        var cancellationThread = new Thread(() => {
            encodingStarted.Wait();
            cancellation.Cancel();
            cancellationRequested.Set();
        }) {
            IsBackground = true,
            Name = "OfficeIMO release-quality encoder cancellation"
        };
        cancellationThread.Start();
        int checkpointCount = 0;
        try {
            using var output = new MemoryStream();
            ExpectException<OperationCanceledException>(() =>
                OfficeRasterImageEncoder.EncodeTo(
                    _source,
                    _format,
                    output,
                    _options,
                    maximumEncodedBytes: long.MaxValue,
                    cancellationToken: cancellation.Token,
                    checkpointObserver: checkpoint => {
                        if (checkpoint != expectedCheckpoint || Interlocked.Increment(ref checkpointCount) != 2) return;
                        encodingStarted.Set();
                        if (!cancellationRequested.Wait(TimeSpan.FromSeconds(5))) {
                            throw new InvalidOperationException("The synchronized cancellation request did not arrive.");
                        }
                    }));
        } finally {
            encodingStarted.Set();
            cancellationThread.Join();
        }
        if (checkpointCount < 2) {
            throw new InvalidOperationException($"{ScenarioId} {Format} did not progress through two compression checkpoints.");
        }
    }

    private static TException ExpectException<TException>(Action action) where TException : Exception {
        try {
            action();
        } catch (TException exception) {
            return exception;
        }
        throw new InvalidOperationException($"Expected {typeof(TException).Name}.");
    }

    private static double CalculateMeanAbsoluteRgbError(OfficeRasterImage expected, OfficeRasterImage actual) {
        byte[] expectedPixels = expected.GetPixels();
        byte[] actualPixels = actual.GetPixels();
        long total = 0L;
        long channels = 0L;
        for (int i = 0; i < expectedPixels.Length; i += 4) {
            total += Math.Abs(expectedPixels[i] - actualPixels[i]);
            total += Math.Abs(expectedPixels[i + 1] - actualPixels[i + 1]);
            total += Math.Abs(expectedPixels[i + 2] - actualPixels[i + 2]);
            channels += 3L;
        }
        return channels == 0L ? 0D : total / (double)channels;
    }
}
