using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Source-neutral sparkline rendering style used by raster and SVG Drawing renderers.
/// </summary>
public sealed class OfficeSparklineStyle {
    /// <summary>Default line and fallback point color.</summary>
    public OfficeColor SeriesColor { get; set; } = OfficeColor.FromRgb(37, 99, 235);

    /// <summary>Axis color used when <see cref="DisplayAxis"/> is enabled and values cross zero.</summary>
    public OfficeColor AxisColor { get; set; } = OfficeColor.FromRgb(128, 128, 128);

    /// <summary>Whether to draw a zero axis when the data range crosses zero.</summary>
    public bool DisplayAxis { get; set; }

    /// <summary>Padding inside the sparkline rectangle, in rendered units.</summary>
    public double Padding { get; set; } = 3D;

    /// <summary>Line stroke width, in rendered units.</summary>
    public double LineStrokeWidth { get; set; } = 1.35D;

    /// <summary>Axis stroke width, in rendered units.</summary>
    public double AxisStrokeWidth { get; set; } = 1D;

    /// <summary>Axis horizontal inset from the outer sparkline rectangle, in rendered units.</summary>
    public double AxisInset { get; set; } = 2D;

    /// <summary>Marker diameter for line sparkline markers, in rendered units.</summary>
    public double MarkerDiameter { get; set; } = 3D;

    /// <summary>Column bar width as a ratio of the per-point slot width.</summary>
    public double ColumnWidthRatio { get; set; } = 0.62D;

    /// <summary>Win/loss bar height as a ratio of the plot height.</summary>
    public double WinLossHeightRatio { get; set; } = 0.42D;

    /// <summary>Optional per-point colors and marker flags resolved by the source adapter.</summary>
    public IReadOnlyList<OfficeSparklinePointStyle>? PointStyles { get; set; }
}
