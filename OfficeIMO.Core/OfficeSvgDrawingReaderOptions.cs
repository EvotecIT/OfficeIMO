namespace OfficeIMO.Drawing;

/// <summary>
/// Controls bounded SVG import limits for trusted inputs that legitimately contain many elements.
/// </summary>
public sealed class OfficeSvgDrawingReaderOptions {
    /// <summary>
    /// Font family used by SVG text that does not declare one. The default preserves
    /// the existing shared-drawing contract; browser-oriented hosts can select a
    /// generic family that matches their user-agent profile.
    /// </summary>
    public string DefaultFontFamily { get; set; } = "Arial";

    /// <summary>
    /// Scoped font faces available when SVG text must be converted to painted vector outlines.
    /// Callers can register web fonts without changing process-wide font state.
    /// </summary>
    public OfficeFontFaceCollection Fonts { get; } = new OfficeFontFaceCollection();

    /// <summary>
    /// Optional caller-owned renderer for bounded inline XHTML inside SVG <c>foreignObject</c> elements.
    /// The returned drawing must use the exact viewport dimensions supplied by the context.
    /// </summary>
    public OfficeSvgForeignObjectRenderer? ForeignObjectRenderer { get; set; }

    /// <summary>
    /// Width of a caller-resolved SVG viewport, in CSS pixels. Supply with
    /// <see cref="ViewportHeight"/> when a host layout owns the viewport size.
    /// </summary>
    public double? ViewportWidth { get; set; }

    /// <summary>
    /// Height of a caller-resolved SVG viewport, in CSS pixels. Supply with
    /// <see cref="ViewportWidth"/> when a host layout owns the viewport size.
    /// </summary>
    public double? ViewportHeight { get; set; }

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

    /// <summary>Default maximum number of full SVG visual comparisons performed by content-safety inspection.</summary>
    public const int DefaultMaximumContentSafetyVisualComparisons = 32;

    /// <summary>Hard maximum number of SVG visual comparisons accepted from a caller.</summary>
    public const int MaximumAllowedContentSafetyVisualComparisons = 128;

    /// <summary>Default cumulative rendered-pixel budget for SVG content-safety inspection.</summary>
    public const long DefaultMaximumContentSafetyVisualPixels = 32L * 1024L * 1024L;

    /// <summary>Hard maximum cumulative rendered-pixel budget accepted from a caller.</summary>
    public const long MaximumAllowedContentSafetyVisualPixels = 128L * 1024L * 1024L;

    /// <summary>
    /// Maximum number of descendant and expanded reference elements. Increase this only for trusted SVG input.
    /// </summary>
    public int MaximumElements { get; set; } = DefaultMaximumElements;

    /// <summary>Maximum SVG viewport width or height. Increase this only for trusted SVG input.</summary>
    public double MaximumViewportDimension { get; set; } = DefaultMaximumViewportDimension;

    /// <summary>Maximum SVG viewport width-times-height area. Increase this only for trusted SVG input.</summary>
    public double MaximumViewportPixels { get; set; } = DefaultMaximumViewportPixels;

    /// <summary>
    /// Maximum full-document visual comparisons used for paint-order and background evidence.
    /// Set this to zero to use structural inspection only; the report records that visual comparison was disabled.
    /// </summary>
    public int MaximumContentSafetyVisualComparisons { get; set; } = DefaultMaximumContentSafetyVisualComparisons;

    /// <summary>Cumulative rendered-pixel work allowed across the SVG content-safety baseline and comparison renders.</summary>
    public long MaximumContentSafetyVisualPixels { get; set; } = DefaultMaximumContentSafetyVisualPixels;
}
