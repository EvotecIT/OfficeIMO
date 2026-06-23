namespace OfficeIMO.Drawing;

/// <summary>
/// Preferred side for chart axis tick labels.
/// </summary>
public enum OfficeChartAxisTickLabelPosition {
    /// <summary>Render labels beside the default axis side.</summary>
    NextTo,

    /// <summary>Render labels on the low-value or low-category side of the plot area.</summary>
    Low,

    /// <summary>Render labels on the high-value or high-category side of the plot area.</summary>
    High,

    /// <summary>Do not render tick labels.</summary>
    None
}
