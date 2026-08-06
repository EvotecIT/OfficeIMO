namespace OfficeIMO.Drawing;

/// <summary>
/// Deterministic text measurement result in pixels.
/// </summary>
public readonly struct OfficeTextMetrics {
    internal OfficeTextMetrics(double widthPixels, double lineHeightPixels, double spaceWidthPixels, double maximumDigitWidthPixels) {
        WidthPixels = widthPixels;
        LineHeightPixels = lineHeightPixels;
        SpaceWidthPixels = spaceWidthPixels;
        MaximumDigitWidthPixels = maximumDigitWidthPixels;
    }

    /// <summary>Estimated text width in pixels.</summary>
    public double WidthPixels { get; }

    /// <summary>Estimated single-line height in pixels.</summary>
    public double LineHeightPixels { get; }

    /// <summary>Estimated width of a space character in pixels.</summary>
    public double SpaceWidthPixels { get; }

    /// <summary>Estimated maximum digit width in pixels.</summary>
    public double MaximumDigitWidthPixels { get; }
}
