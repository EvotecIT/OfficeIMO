using OfficeIMO.Drawing;

namespace OfficeIMO.OneNote;

/// <summary>Controls dependency-free OneNote page layout, rendering, and image export.</summary>
public class OneNotePageRenderingOptions : OfficeImageExportOptions {
    /// <summary>Whether the page title is rendered.</summary>
    public bool IncludeTitle { get; set; } = true;

    /// <summary>Whether embedded pictures, including printout backgrounds, are rendered.</summary>
    public bool IncludeImages { get; set; } = true;

    /// <summary>Whether native ink strokes are rendered.</summary>
    public bool IncludeInk { get; set; } = true;

    /// <summary>Whether structured mathematical expressions are typeset.</summary>
    public bool IncludeMath { get; set; } = true;

    /// <summary>Whether attachment and recording placeholders are rendered.</summary>
    public bool IncludeAttachmentPlaceholders { get; set; } = true;

    /// <summary>Maximum bytes materialized from any single lazy image payload.</summary>
    public long MaxImageBytes { get; set; } = 64L * 1024L * 1024L;

    /// <summary>Maximum decoded pixels allocated by one raster page export.</summary>
    public long MaximumRasterPixels { get; set; } = 100_000_000L;

    /// <summary>Optional decoder for source image formats not handled by the dependency-free Drawing core.</summary>
    public IOfficeRasterImageCodec? ImageCodec { get; set; }

    /// <summary>Minimum width used for automatically sized pages, in points.</summary>
    public double AutomaticPageWidthPoints { get; set; } = 612D;

    /// <summary>Minimum height used for automatically sized pages, in points.</summary>
    public double AutomaticPageHeightPoints { get; set; } = 792D;

    /// <summary>Extra space retained beyond inferred content bounds, in points.</summary>
    public double AutomaticPagePaddingPoints { get; set; } = 36D;

    /// <summary>Default body font used when a OneNote run does not name one.</summary>
    public OfficeFontInfo DefaultFont { get; set; } = new OfficeFontInfo("Calibri", 11D);

    /// <summary>Reusable ink-rendering settings.</summary>
    public OfficeInkRenderOptions Ink { get; set; } = new OfficeInkRenderOptions();

    /// <summary>Reusable mathematical-rendering settings.</summary>
    public OfficeMathRenderOptions Math { get; set; } = new OfficeMathRenderOptions();

    /// <summary>Creates a detached copy.</summary>
    public OneNotePageRenderingOptions Clone() => new OneNotePageRenderingOptions {
        Scale = Scale,
        BackgroundColor = BackgroundColor,
        RasterEncoding = RasterEncoding?.Clone() ?? new OfficeRasterEncodingOptions(),
        IncludeTitle = IncludeTitle,
        IncludeImages = IncludeImages,
        IncludeInk = IncludeInk,
        IncludeMath = IncludeMath,
        IncludeAttachmentPlaceholders = IncludeAttachmentPlaceholders,
        MaxImageBytes = MaxImageBytes,
        MaximumRasterPixels = MaximumRasterPixels,
        ImageCodec = ImageCodec,
        AutomaticPageWidthPoints = AutomaticPageWidthPoints,
        AutomaticPageHeightPoints = AutomaticPageHeightPoints,
        AutomaticPagePaddingPoints = AutomaticPagePaddingPoints,
        DefaultFont = DefaultFont,
        Ink = Ink?.Clone() ?? new OfficeInkRenderOptions(),
        Math = Math?.Clone() ?? new OfficeMathRenderOptions()
    };

    internal void Validate() {
        ValidateScale(Scale);
        if (MaxImageBytes < 1) throw new ArgumentOutOfRangeException(nameof(MaxImageBytes));
        if (MaximumRasterPixels < 1L) throw new ArgumentOutOfRangeException(nameof(MaximumRasterPixels));
        ValidatePositive(AutomaticPageWidthPoints, nameof(AutomaticPageWidthPoints));
        ValidatePositive(AutomaticPageHeightPoints, nameof(AutomaticPageHeightPoints));
        if (double.IsNaN(AutomaticPagePaddingPoints) || double.IsInfinity(AutomaticPagePaddingPoints) || AutomaticPagePaddingPoints < 0D) {
            throw new ArgumentOutOfRangeException(nameof(AutomaticPagePaddingPoints));
        }
        if (DefaultFont.Size <= 0D || double.IsNaN(DefaultFont.Size) || double.IsInfinity(DefaultFont.Size)) {
            throw new ArgumentOutOfRangeException(nameof(DefaultFont));
        }
        if (Ink == null) throw new InvalidOperationException("Ink rendering options cannot be null.");
        if (Math == null) throw new InvalidOperationException("Math rendering options cannot be null.");
    }

    private static void ValidatePositive(double value, string name) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) throw new ArgumentOutOfRangeException(name);
    }
}
