namespace OfficeIMO.Drawing;

/// <summary>
/// Controls bounded SVG import limits for trusted inputs that legitimately contain many elements.
/// </summary>
public sealed class OfficeSvgDrawingReaderOptions {
    /// <summary>Default maximum number of descendant and expanded reference elements.</summary>
    public const int DefaultMaximumElements = 10000;

    /// <summary>Hard maximum accepted by the reader, even when explicitly requested.</summary>
    public const int MaximumAllowedElements = 100000;

    /// <summary>Default maximum width or height of an imported SVG viewport.</summary>
    public const double DefaultMaximumViewportDimension = 8192D;

    /// <summary>Hard maximum viewport dimension accepted for explicitly trusted SVG input.</summary>
    public const double MaximumAllowedViewportDimension = 1000000D;

    /// <summary>Default maximum viewport area accepted by the reader.</summary>
    public const double DefaultMaximumViewportPixels = 16D * 1024D * 1024D;

    /// <summary>Hard maximum viewport area accepted for explicitly trusted SVG input.</summary>
    public const double MaximumAllowedViewportPixels = 256D * 1024D * 1024D;

    /// <summary>
    /// Maximum number of descendant and expanded reference elements. Increase this only for trusted SVG input.
    /// </summary>
    public int MaximumElements { get; set; } = DefaultMaximumElements;

    /// <summary>Maximum SVG viewport width or height. Increase this only for trusted SVG input.</summary>
    public double MaximumViewportDimension { get; set; } = DefaultMaximumViewportDimension;

    /// <summary>Maximum SVG viewport width-times-height area. Increase this only for trusted SVG input.</summary>
    public double MaximumViewportPixels { get; set; } = DefaultMaximumViewportPixels;
}
