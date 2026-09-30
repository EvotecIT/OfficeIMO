using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Resource policy for rendered page alignment.</summary>
public sealed class PdfPageChangeOptions {
    /// <summary>Raster scale used to find same-looking pages. Higher values see smaller visual changes.</summary>
    public double RenderScale { get; set; } = 1D;
    /// <summary>Background used for alpha compositing.</summary>
    public OfficeColor Background { get; set; } = OfficeColor.White;
    /// <summary>Maximum pages in each document.</summary>
    public int MaxPagesPerDocument { get; set; } = 100;
    /// <summary>Maximum raster pixels for one page.</summary>
    public long MaxPixelsPerPage { get; set; } = 20_000_000L;
    /// <summary>Maximum raster pixels across both documents.</summary>
    public long MaxTotalPixels { get; set; } = 100_000_000L;

    internal void Validate() {
        if (RenderScale <= 0D || double.IsNaN(RenderScale) || double.IsInfinity(RenderScale)) throw new ArgumentOutOfRangeException(nameof(RenderScale));
        if (MaxPagesPerDocument <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPagesPerDocument));
        if (MaxPixelsPerPage <= 0) throw new ArgumentOutOfRangeException(nameof(MaxPixelsPerPage));
        if (MaxTotalPixels <= 0) throw new ArgumentOutOfRangeException(nameof(MaxTotalPixels));
    }
}
