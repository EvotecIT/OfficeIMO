namespace OfficeIMO.Drawing;

/// <summary>
/// Physical side where a chart axis crosses the perpendicular axis.
/// </summary>
public enum OfficeChartAxisCrossingPosition {
    /// <summary>Use the renderer's default zero or low-side crossing.</summary>
    AutoZero,

    /// <summary>Cross at the high-value or high-category side of the plot area.</summary>
    Maximum
}
