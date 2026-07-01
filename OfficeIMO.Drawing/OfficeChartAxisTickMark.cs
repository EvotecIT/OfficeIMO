namespace OfficeIMO.Drawing;

/// <summary>
/// Describes simple chart axis tick mark placement for dependency-free chart rendering.
/// </summary>
public enum OfficeChartAxisTickMark {
    /// <summary>No tick marks are rendered.</summary>
    None,

    /// <summary>Tick marks are rendered inside the plot area.</summary>
    Inside,

    /// <summary>Tick marks are rendered outside the plot area.</summary>
    Outside,

    /// <summary>Tick marks cross the axis line.</summary>
    Cross
}
