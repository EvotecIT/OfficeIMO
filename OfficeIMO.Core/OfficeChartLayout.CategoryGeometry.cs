using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeChartLayout {
    /// <summary>The gap between bar clusters as a percentage of one bar's width, from 0 to 500. New charts use the shared theme; imported Office charts without a value use 150.</summary>
    public int BarGapWidthPercent { get; private set; } = OfficeTheme.Default.ChartStyle.BarGapPercent;

    /// <summary>How far bars in a cluster overlap, from -100 (a full bar apart) to 100 (fully overlapped).</summary>
    /// <remarks>Stacked bars always overlap fully; this value applies to clustered bars.</remarks>
    public int BarOverlapPercent { get; private set; }

    /// <summary>Whether line and area points sit on the category tick marks (OOXML <c>midCat</c>)
    /// rather than in the middle of each category (OOXML <c>between</c>, the default).</summary>
    public bool CategoryPointsOnTickMarks { get; private set; }

    /// <summary>Returns a layout copy with the requested bar gap width and overlap.</summary>
    /// <param name="gapWidthPercent">The gap between clusters as a percentage of one bar's width, from 0 to 500.</param>
    /// <param name="overlapPercent">The overlap of clustered bars, from -100 to 100.</param>
    public OfficeChartLayout WithBarGeometry(int gapWidthPercent, int overlapPercent) {
        if (gapWidthPercent < 0 || gapWidthPercent > 500) throw new ArgumentOutOfRangeException(nameof(gapWidthPercent));
        if (overlapPercent < -100 || overlapPercent > 100) throw new ArgumentOutOfRangeException(nameof(overlapPercent));
        var copy = (OfficeChartLayout)MemberwiseClone();
        copy.BarGapWidthPercent = gapWidthPercent;
        copy.BarOverlapPercent = overlapPercent;
        return copy;
    }

    /// <summary>Returns a layout copy that places line and area points on the category tick marks or between them.</summary>
    public OfficeChartLayout WithCategoryPointsOnTickMarks(bool onTickMarks) {
        var copy = (OfficeChartLayout)MemberwiseClone();
        copy.CategoryPointsOnTickMarks = onTickMarks;
        return copy;
    }

    /// <summary>Whether a vertical line joins the highest and lowest line-series value in each category.</summary>
    public bool ShowHighLowLines { get; private set; }

    /// <summary>The high-low line colour, or null for the chart text colour.</summary>
    public OfficeColor? HighLowLineColor { get; private set; }

    /// <summary>The high-low line width in points, or null for 0.75 points.</summary>
    public double? HighLowLineWidth { get; private set; }

    /// <summary>Returns a layout copy that draws or hides high-low lines across the line series of each category.</summary>
    /// <param name="show">Whether to draw high-low lines.</param>
    /// <param name="color">The line colour, or null for the chart text colour.</param>
    /// <param name="width">The line width in points, greater than zero and at most 1584, or null for 0.75 points.</param>
    public OfficeChartLayout WithHighLowLines(bool show, OfficeColor? color = null, double? width = null) {
        if (width.HasValue && (double.IsNaN(width.Value) || width.Value <= 0D || width.Value > OfficeChartStyleBounds.MaximumLineWidthPoints))
            throw new ArgumentOutOfRangeException(nameof(width));
        var copy = (OfficeChartLayout)MemberwiseClone();
        copy.ShowHighLowLines = show;
        copy.HighLowLineColor = color;
        copy.HighLowLineWidth = width;
        return copy;
    }
}
