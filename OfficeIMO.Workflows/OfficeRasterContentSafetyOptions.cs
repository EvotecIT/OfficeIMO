using OfficeIMO.ContentSafety;
using OfficeIMO.Drawing;

namespace OfficeIMO.Workflows;

/// <summary>Bounds OCR-backed concealment inspection and optional raster-region redaction.</summary>
public sealed class OfficeRasterContentSafetyOptions {
    /// <summary>Shared content-safety thresholds and finding limits.</summary>
    public OfficeContentSafetyOptions Inspection { get; set; } = new();

    /// <summary>Maximum decoded pixels retained for one static image. Defaults to 16 million.</summary>
    public long MaximumDecodedPixels { get; set; } = 16_000_000L;

    /// <summary>Maximum OCR spans accepted from the provider. Defaults to 4,096.</summary>
    public int MaximumOcrSpans { get; set; } = 4096;

    /// <summary>
    /// Maximum cumulative region-pixel work accepted during inspection or selected-region redaction.
    /// Defaults to 64 million.
    /// </summary>
    public long MaximumPixelAnalysisWork { get; set; } = 64_000_000L;

    /// <summary>
    /// Maximum cumulative OCR-region intersection comparisons during selected-region redaction.
    /// Defaults to one million.
    /// </summary>
    public long MaximumRegionComparisons { get; set; } = 1_000_000L;

    /// <summary>Pixel height at or below which a recognized span is treated as tiny. Defaults to three pixels.</summary>
    public int MaximumTinyTextHeightPixels { get; set; } = 3;

    /// <summary>
    /// Maximum alpha channel value that proves an entire OCR region is nearly transparent.
    /// Defaults to 16 out of 255.
    /// </summary>
    public byte MaximumConcealedAlpha { get; set; } = 16;

    /// <summary>Total timeout for one OCR provider call. Defaults to 30 seconds.</summary>
    public TimeSpan OcrTimeout { get; set; } = TimeSpan.FromSeconds(30);

    /// <summary>
    /// Enables explicit opaque-rectangle redaction for findings with bounded geometry and sufficient confidence.
    /// Disabled by default; disabled findings remain report-only.
    /// </summary>
    public bool EnableOpaqueRectangleRedaction { get; set; }

    /// <summary>Minimum provider confidence required before a region can be redacted. Defaults to 0.8.</summary>
    public double MinimumOcrConfidenceForRedaction { get; set; } = 0.8D;

    /// <summary>Pixels added around each selected OCR region during redaction. Defaults to one.</summary>
    public int RedactionPaddingPixels { get; set; } = 1;

    /// <summary>Opaque color written over selected OCR regions. Defaults to black.</summary>
    public OfficeColor RedactionColor { get; set; } = OfficeColor.Black;

    /// <summary>Maximum normalized PNG bytes produced for OCR or redaction output. Defaults to 128 MiB.</summary>
    public long MaximumOutputBytes { get; set; } = 128L * 1024L * 1024L;

    internal Snapshot Capture() {
        if (Inspection == null) throw new ArgumentNullException(nameof(Inspection));
        var inspection = new OfficeContentSafetyOptions {
            MaxInputBytes = Inspection.MaxInputBytes,
            MaxPackageEntries = Inspection.MaxPackageEntries,
            MaxExpandedPackageBytes = Inspection.MaxExpandedPackageBytes,
            MaxCharacters = Inspection.MaxCharacters,
            MaxFindings = Inspection.MaxFindings,
            MaxPreviewCharacters = Inspection.MaxPreviewCharacters,
            MaximumTinyFontSizePoints = Inspection.MaximumTinyFontSizePoints,
            MinimumVisibleContrastRatio = Inspection.MinimumVisibleContrastRatio,
            IncludeNonPrimaryContent = Inspection.IncludeNonPrimaryContent,
            DetectInstructionLikeText = Inspection.DetectInstructionLikeText,
            IncludeTextIntegrityEvidence = Inspection.IncludeTextIntegrityEvidence
        };
        if (inspection.MaxInputBytes <= 0L || inspection.MaxInputBytes > int.MaxValue) {
            throw new ArgumentOutOfRangeException(nameof(Inspection));
        }
        if (MaximumDecodedPixels <= 0L || MaximumDecodedPixels > 50_000_000L) {
            throw new ArgumentOutOfRangeException(nameof(MaximumDecodedPixels));
        }
        if (MaximumOcrSpans <= 0 || MaximumOcrSpans > 100_000) {
            throw new ArgumentOutOfRangeException(nameof(MaximumOcrSpans));
        }
        if (MaximumPixelAnalysisWork <= 0L || MaximumPixelAnalysisWork > 1_000_000_000L) {
            throw new ArgumentOutOfRangeException(nameof(MaximumPixelAnalysisWork));
        }
        if (MaximumRegionComparisons <= 0L || MaximumRegionComparisons > 100_000_000L) {
            throw new ArgumentOutOfRangeException(nameof(MaximumRegionComparisons));
        }
        if (MaximumTinyTextHeightPixels < 0 || MaximumTinyTextHeightPixels > 1024) {
            throw new ArgumentOutOfRangeException(nameof(MaximumTinyTextHeightPixels));
        }
        if (OcrTimeout <= TimeSpan.Zero || OcrTimeout > TimeSpan.FromHours(1)) {
            throw new ArgumentOutOfRangeException(nameof(OcrTimeout));
        }
        if (double.IsNaN(MinimumOcrConfidenceForRedaction) ||
            double.IsInfinity(MinimumOcrConfidenceForRedaction) ||
            MinimumOcrConfidenceForRedaction < 0D || MinimumOcrConfidenceForRedaction > 1D) {
            throw new ArgumentOutOfRangeException(nameof(MinimumOcrConfidenceForRedaction));
        }
        if (RedactionPaddingPixels < 0 || RedactionPaddingPixels > 1024) {
            throw new ArgumentOutOfRangeException(nameof(RedactionPaddingPixels));
        }
        if (RedactionColor.A != byte.MaxValue) {
            throw new ArgumentException("Raster redaction requires a fully opaque color.", nameof(RedactionColor));
        }
        if (MaximumOutputBytes <= 0L || MaximumOutputBytes > 128L * 1024L * 1024L) {
            throw new ArgumentOutOfRangeException(nameof(MaximumOutputBytes));
        }
        _ = new OfficeContentSafetyBuilder("Raster Image", inspection);
        return new Snapshot(
            inspection,
            MaximumDecodedPixels,
            MaximumOcrSpans,
            MaximumPixelAnalysisWork,
            MaximumRegionComparisons,
            MaximumTinyTextHeightPixels,
            MaximumConcealedAlpha,
            OcrTimeout,
            EnableOpaqueRectangleRedaction,
            MinimumOcrConfidenceForRedaction,
            RedactionPaddingPixels,
            RedactionColor,
            MaximumOutputBytes);
    }

    internal sealed class Snapshot {
        internal Snapshot(
            OfficeContentSafetyOptions inspection,
            long maximumDecodedPixels,
            int maximumOcrSpans,
            long maximumPixelAnalysisWork,
            long maximumRegionComparisons,
            int maximumTinyTextHeightPixels,
            byte maximumConcealedAlpha,
            TimeSpan ocrTimeout,
            bool enableOpaqueRectangleRedaction,
            double minimumOcrConfidenceForRedaction,
            int redactionPaddingPixels,
            OfficeColor redactionColor,
            long maximumOutputBytes) {
            Inspection = inspection;
            MaximumDecodedPixels = maximumDecodedPixels;
            MaximumOcrSpans = maximumOcrSpans;
            MaximumPixelAnalysisWork = maximumPixelAnalysisWork;
            MaximumRegionComparisons = maximumRegionComparisons;
            MaximumTinyTextHeightPixels = maximumTinyTextHeightPixels;
            MaximumConcealedAlpha = maximumConcealedAlpha;
            OcrTimeout = ocrTimeout;
            EnableOpaqueRectangleRedaction = enableOpaqueRectangleRedaction;
            MinimumOcrConfidenceForRedaction = minimumOcrConfidenceForRedaction;
            RedactionPaddingPixels = redactionPaddingPixels;
            RedactionColor = redactionColor;
            MaximumOutputBytes = maximumOutputBytes;
        }

        internal OfficeContentSafetyOptions Inspection { get; }
        internal long MaximumDecodedPixels { get; }
        internal int MaximumOcrSpans { get; }
        internal long MaximumPixelAnalysisWork { get; }
        internal long MaximumRegionComparisons { get; }
        internal int MaximumTinyTextHeightPixels { get; }
        internal byte MaximumConcealedAlpha { get; }
        internal TimeSpan OcrTimeout { get; }
        internal bool EnableOpaqueRectangleRedaction { get; }
        internal double MinimumOcrConfidenceForRedaction { get; }
        internal int RedactionPaddingPixels { get; }
        internal OfficeColor RedactionColor { get; }
        internal long MaximumOutputBytes { get; }
    }
}