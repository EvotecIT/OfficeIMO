using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

/// <summary>Controls the shared-Drawing layout diagnostics projected over one PDF page.</summary>
public sealed class PdfLayoutDebugOverlayOptions {
    /// <summary>Default maximum number of pixels allocated by one PNG overlay.</summary>
    public const long DefaultMaxRasterPixels = 64_000_000;

    /// <summary>Draw boxes around whitespace-delimited words.</summary>
    public bool ShowWords { get; set; } = true;

    /// <summary>Draw boxes around inferred text lines.</summary>
    public bool ShowLines { get; set; } = true;

    /// <summary>Draw boxes around inferred paragraph regions.</summary>
    public bool ShowRegions { get; set; } = true;

    /// <summary>Draw reading-order numbers for lines.</summary>
    public bool ShowReadingOrder { get; set; } = true;

    /// <summary>Maximum boxes and labels added to one overlay.</summary>
    public int MaxElements { get; set; } = 20000;

    /// <summary>Maximum number of pixels allocated when rasterizing one overlay.</summary>
    public long MaxRasterPixels { get; set; } = DefaultMaxRasterPixels;

    /// <summary>Word-box color.</summary>
    public OfficeColor WordColor { get; set; } = OfficeColor.FromRgb(40, 120, 220);

    /// <summary>Line-box color.</summary>
    public OfficeColor LineColor { get; set; } = OfficeColor.FromRgb(30, 170, 90);

    /// <summary>Paragraph-region color.</summary>
    public OfficeColor RegionColor { get; set; } = OfficeColor.FromRgb(220, 80, 60);
}
