using OfficeIMO.ContentSafety;
using OfficeIMO.Drawing;
using OfficeIMO.Ocr;

namespace OfficeIMO.Workflows;

/// <summary>OCR-backed concealed-text inspection and explicit rectangular redaction for bounded raster images.</summary>
public static partial class OfficeRasterContentSafety {
    private const string ReportFormat = "Raster Image";
    private const string NormalizedMediaType = "image/png";

    /// <summary>
    /// Inspects a single-frame raster image using caller-owned OCR and pixel evidence.
    /// Metadata and provider text without bounded geometry never become concealment findings.
    /// </summary>
    public static async Task<OfficeContentSafetyReport> InspectAsync(
        byte[] imageBytes,
        IOcrEngine engine,
        OfficeRasterContentSafetyOptions? options = null,
        CancellationToken cancellationToken = default) {
        ArgumentNullException.ThrowIfNull(imageBytes);
        ArgumentNullException.ThrowIfNull(engine);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeRasterContentSafetyOptions.Snapshot snapshot =
            (options ?? new OfficeRasterContentSafetyOptions()).Capture();
        OfficeContentSafetyInputGuard.ValidateBytes(imageBytes, snapshot.Inspection);
        byte[] input = (byte[])imageBytes.Clone();
        OcrEngineExecution execution = OcrEngineRunner.CreateExecution(engine);
        var budget = new RasterWorkBudget(snapshot);
        AnalysisState state = await InspectCoreAsync(
                input,
                execution,
                snapshot,
                budget,
                imageBytes.LongLength + 24L,
                cancellationToken)
            .ConfigureAwait(false);
        return state.Report;
    }

    /// <summary>Reads and inspects one bounded raster file.</summary>
    public static async Task<OfficeContentSafetyReport> InspectAsync(
        string filePath,
        IOcrEngine engine,
        OfficeRasterContentSafetyOptions? options = null,
        CancellationToken cancellationToken = default) {
        if (string.IsNullOrWhiteSpace(filePath)) throw new ArgumentException("An image path is required.", nameof(filePath));
        ArgumentNullException.ThrowIfNull(engine);
        cancellationToken.ThrowIfCancellationRequested();
        OfficeRasterContentSafetyOptions effective = options ?? new OfficeRasterContentSafetyOptions();
        OfficeRasterContentSafetyOptions.Snapshot snapshot = effective.Capture();
        byte[] input = OfficeContentSafetyInputGuard.ReadAllBytes(
            filePath,
            snapshot.Inspection,
            inspectZipPackage: false,
            cancellationToken: cancellationToken);
        OcrEngineExecution execution = OcrEngineRunner.CreateExecution(engine);
        var budget = new RasterWorkBudget(snapshot);
        AnalysisState state = await InspectCoreAsync(
                input,
                execution,
                snapshot,
                budget,
                additionalRetainedManagedBytes: 0L,
                cancellationToken)
            .ConfigureAwait(false);
        return state.Report;
    }

    private static async Task<AnalysisState> InspectCoreAsync(
        byte[] input,
        OcrEngineExecution execution,
        OfficeRasterContentSafetyOptions.Snapshot options,
        RasterWorkBudget budget,
        long additionalRetainedManagedBytes,
        CancellationToken cancellationToken) {
        cancellationToken.ThrowIfCancellationRequested();
        OfficeContentSafetyInputGuard.ValidateBytes(input, options.Inspection);
        if (!execution.Capabilities.SupportsMediaType(NormalizedMediaType)) {
            throw new NotSupportedException("The configured OCR engine does not advertise support for normalized PNG input.");
        }
        var decodeOptions = new OfficeRasterDecodeOptions {
            MaximumEncodedBytes = checked((int)Math.Min(options.Inspection.MaxInputBytes, 128L * 1024L * 1024L)),
            MaximumDecodedPixels = options.MaximumDecodedPixels,
            FrameLossPolicy = OfficeRasterFrameLossPolicy.RejectMultipleFrames,
            CancellationToken = cancellationToken,
            RetainedManagedBytes = additionalRetainedManagedBytes
        };
        if (!OfficeRasterImageDecoder.TryDecode(input, decodeOptions, out OfficeRasterImage? image, out OfficeRasterDecodeInfo decodeInfo) || image == null) {
            throw new InvalidDataException(decodeInfo.Diagnostic ??
                "The raster image could not be decoded within the configured limits.");
        }

        OfficeImageMetadataSnapshot metadata = OfficeImageMetadataInspector.Inspect(
            input,
            decodeInfo.Format,
            checked(input.LongLength + image.PixelBuffer.LongLength + 48L + additionalRetainedManagedBytes),
            cancellationToken);
        if (metadata.HasColorRenderingMetadata) {
            throw new InvalidDataException(
                "Raster content-safety inspection rejects color profiles and non-sRGB PNG color-rendering metadata because the managed decoder does not color-normalize them to sRGB.");
        }
        if ((metadata.Kinds & OfficeImageMetadataKinds.Orientation) != 0 &&
            decodeInfo.Format != OfficeImageFormat.Jpeg && decodeInfo.Format != OfficeImageFormat.Tiff) {
            throw new InvalidDataException(
                "Raster content-safety inspection rejects embedded orientation that the managed decoder cannot visibly normalize.");
        }

        byte[] normalized = OfficeRasterImageEncoder.Encode(
            image,
            OfficeImageExportFormat.Png,
            CreateMetadataFreePngEncodingOptions(),
            options.MaximumOutputBytes,
            cancellationToken,
            checked(input.LongLength + 24L + additionalRetainedManagedBytes));
        var request = new OcrRequest {
            Operation = OcrOperation.RecognizeText,
            Payload = normalized,
            MediaType = NormalizedMediaType,
            FileName = "officeimo-content-safety.png",
            SourceId = "raster-content-safety",
            SourceName = "normalized-raster",
            CandidateId = "frame-1",
            CandidateKind = "raster-frame",
            PageNumber = 1,
            PixelWidth = image.Width,
            PixelHeight = image.Height,
            Region = new OcrRegion { X = 0D, Y = 0D, Width = image.Width, Height = image.Height },
            RegionCoordinateUnit = OcrCoordinateUnit.Pixels
        };
        OcrResult result = await execution.RecognizeAsync(request, options.OcrTimeout, cancellationToken)
            .ConfigureAwait(false);
        cancellationToken.ThrowIfCancellationRequested();
        return Analyze(image, result, execution.Id, options, budget, cancellationToken);
    }

    private static OfficeRasterEncodingOptions CreateMetadataFreePngEncodingOptions() => new() {
        WriteResolutionMetadata = false
    };

    private sealed class AnalysisState {
        internal AnalysisState(
            OfficeRasterImage image,
            OfficeContentSafetyReport report,
            IReadOnlyDictionary<string, RasterTarget> targets,
            IReadOnlyList<RasterTarget> recognizedTargets) {
            Image = image;
            Report = report;
            Targets = targets;
            RecognizedTargets = recognizedTargets;
        }

        internal OfficeRasterImage Image { get; }
        internal OfficeContentSafetyReport Report { get; }
        internal IReadOnlyDictionary<string, RasterTarget> Targets { get; }
        internal IReadOnlyList<RasterTarget> RecognizedTargets { get; }
    }

    private sealed class RasterTarget {
        internal RasterTarget(OcrTextSpan span, PixelRegion region) {
            Region = region;
            Level = span.Level;
            LineId = span.LineId;
            Sequence = span.Sequence;
            Text = span.Text ?? string.Empty;
            Confidence = span.Confidence;
        }

        internal PixelRegion Region { get; }
        internal OcrTextSpanLevel Level { get; }
        internal string? LineId { get; }
        internal int Sequence { get; }
        internal string Text { get; }
        internal double? Confidence { get; }
    }

    private sealed class RasterWorkBudget {
        private long _remainingPixels;
        private long _remainingComparisons;
        private int _comparisonCount;

        internal RasterWorkBudget(OfficeRasterContentSafetyOptions.Snapshot options) {
            _remainingPixels = options.MaximumPixelAnalysisWork;
            _remainingComparisons = options.MaximumRegionComparisons;
        }

        internal void ChargePixels(long count) {
            if (count < 0L || count > _remainingPixels) {
                throw new InvalidDataException("Raster pixel-analysis work exceeds the configured cumulative limit.");
            }
            _remainingPixels -= count;
        }

        internal void ChargeComparison(CancellationToken cancellationToken) {
            if (_remainingComparisons <= 0L) {
                throw new InvalidDataException("Raster OCR geometry comparisons exceed the configured cumulative limit.");
            }
            _remainingComparisons--;
            if ((_comparisonCount++ & 4095) == 0) cancellationToken.ThrowIfCancellationRequested();
        }
    }

    private readonly struct PixelRegion {
        internal PixelRegion(int left, int top, int right, int bottom) {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        internal int Left { get; }
        internal int Top { get; }
        internal int Right { get; }
        internal int Bottom { get; }
        internal int Width => Right - Left;
        internal int Height => Bottom - Top;
        internal long Area => (long)Width * Height;

        internal bool Intersects(PixelRegion other) =>
            Left < other.Right && Right > other.Left && Top < other.Bottom && Bottom > other.Top;

        internal bool Contains(PixelRegion other) =>
            Left <= other.Left && Top <= other.Top && Right >= other.Right && Bottom >= other.Bottom;
    }
}
