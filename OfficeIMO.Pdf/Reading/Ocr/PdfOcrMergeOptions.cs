namespace OfficeIMO.Pdf;

/// <summary>Controls OCR rendering, confidence filtering, and native-text overlap removal.</summary>
public sealed class PdfOcrMergeOptions {
    /// <summary>Pages sent to the OCR provider; null means every page.</summary>
    public PdfPageSelection? Selection { get; set; }
    /// <summary>OCR render DPI.</summary>
    public double Dpi { get; set; } = 150D;
    /// <summary>Minimum accepted provider confidence from 0 through 1.</summary>
    public double MinimumConfidence { get; set; } = 0.5D;
    /// <summary>Overlap ratio at which OCR words duplicating native text are removed.</summary>
    public double NativeTextOverlapThreshold { get; set; } = 0.5D;
    /// <summary>Maximum pages sent to the provider.</summary>
    public int MaxPages { get; set; } = 100;
    /// <summary>Maximum pixels rendered per page.</summary>
    public long MaxPixelsPerPage { get; set; } = 100_000_000L;

    internal void Validate() {
        Guard.Positive(Dpi, nameof(Dpi));
        ValidateRatio(MinimumConfidence, nameof(MinimumConfidence));
        ValidateRatio(NativeTextOverlapThreshold, nameof(NativeTextOverlapThreshold));
        Guard.PositiveInteger(MaxPages, nameof(MaxPages));
        if (MaxPixelsPerPage <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPixelsPerPage));
    }

    private static void ValidateRatio(double value, string name) {
        if (value < 0D || value > 1D || double.IsNaN(value)) throw new ArgumentOutOfRangeException(name);
    }
}
