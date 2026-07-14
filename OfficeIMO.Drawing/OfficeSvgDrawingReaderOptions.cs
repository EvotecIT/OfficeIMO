namespace OfficeIMO.Drawing;

/// <summary>
/// Controls bounded SVG import limits for trusted inputs that legitimately contain many elements.
/// </summary>
public sealed class OfficeSvgDrawingReaderOptions {
    /// <summary>Default maximum number of descendant and expanded reference elements.</summary>
    public const int DefaultMaximumElements = 10000;

    /// <summary>Hard maximum accepted by the reader, even when explicitly requested.</summary>
    public const int MaximumAllowedElements = 100000;

    /// <summary>
    /// Maximum number of descendant and expanded reference elements. Increase this only for trusted SVG input.
    /// </summary>
    public int MaximumElements { get; set; } = DefaultMaximumElements;
}
