using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.Pdf;

/// <summary>Controls fail-closed raster replacement for destructive page cropping.</summary>
public sealed class PdfDestructiveCropOptions {
    /// <summary>Observes cancellation during rendering, rewriting, and output validation.</summary>
    public CancellationToken CancellationToken { get; set; }
    /// <summary>Rasterization DPI. Default: 144.</summary>
    public double Dpi { get; set; } = 144D;
    /// <summary>Opaque replacement background.</summary>
    public OfficeColor Background { get; set; } = OfficeColor.White;
    /// <summary>Maximum pixels per destructively cropped page.</summary>
    public long MaxPixelsPerPage { get; set; } = 100_000_000L;
    /// <summary>Allows documented simplified renderer capabilities such as system-font substitution.</summary>
    public bool AllowSimplifiedRendering { get; set; } = true;
}
