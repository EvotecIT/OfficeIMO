using System;
using System.Collections.Generic;
using OfficeIMO.Drawing;

namespace OfficeIMO;

/// <summary>Chart authoring tokens, separate from legacy renderer overrides.</summary>
public sealed class OfficeThemeChartStyle {
    /// <summary>Gets immutable series colors, which may refer to theme roles.</summary>
    public IReadOnlyList<OfficeColor> SeriesPalette { get; }
    /// <summary>Gets axis label fonts.</summary>
    public OfficeThemeFont AxisFont { get; }
    /// <summary>Gets chart title fonts.</summary>
    public OfficeThemeFont TitleFont { get; }
    /// <summary>Gets axis label size.</summary>
    public OfficeLength AxisFontSize { get; }
    /// <summary>Gets chart title size.</summary>
    public OfficeLength TitleFontSize { get; }
    /// <summary>Gets gridline color.</summary>
    public OfficeColor GridlineColor { get; }
    /// <summary>Gets whether value labels are shown.</summary>
    public bool ShowDataLabels { get; }
    /// <summary>Gets whether the series legend is shown.</summary>
    public bool ShowLegend { get; }
    /// <summary>Gets the category axis line colour; the value axis has no line.</summary>
    public OfficeColor AxisLineColor { get; }
    /// <summary>Gets the axis label colour.</summary>
    public OfficeColor AxisTextColor { get; }
    /// <summary>Gets the gridline stroke width.</summary>
    public OfficeLength GridlineWidth { get; }
    /// <summary>Gets the stroke width of line series.</summary>
    public OfficeLength LineWidth { get; }
    /// <summary>Gets the gap between column and bar groups, as a percentage of the bar width.</summary>
    public int BarGapPercent { get; }
    /// <summary>Gets the marker shape for line series; <see cref="OfficeChartMarkerShape.None"/> shows markers only on series with fewer than three points.</summary>
    public OfficeChartMarkerShape Markers { get; }

    /// <summary>Creates immutable chart authoring tokens with a muted axis, 0.5 pt gridlines, 2 pt lines, a 60 % bar gap and no markers.</summary>
    public OfficeThemeChartStyle(IReadOnlyList<OfficeColor> seriesPalette, OfficeThemeFont axisFont, OfficeThemeFont titleFont, OfficeLength axisFontSize, OfficeLength titleFontSize, OfficeColor gridlineColor, bool showDataLabels, bool showLegend)
        : this(seriesPalette, axisFont, titleFont, axisFontSize, titleFontSize, gridlineColor, showDataLabels, showLegend,
            gridlineColor, OfficeColor.TextMuted, OfficeLength.Points(0.5), OfficeLength.Points(2), 60, OfficeChartMarkerShape.None) {
    }

    /// <summary>Creates immutable chart authoring tokens, including axis, stroke, gap and marker details.</summary>
    public OfficeThemeChartStyle(IReadOnlyList<OfficeColor> seriesPalette, OfficeThemeFont axisFont, OfficeThemeFont titleFont, OfficeLength axisFontSize, OfficeLength titleFontSize, OfficeColor gridlineColor, bool showDataLabels, bool showLegend,
        OfficeColor axisLineColor, OfficeColor axisTextColor, OfficeLength gridlineWidth, OfficeLength lineWidth, int barGapPercent, OfficeChartMarkerShape markers) {
        SeriesPalette = ThemeValues.Colors(seriesPalette, nameof(seriesPalette));
        AxisFont = axisFont ?? throw new ArgumentNullException(nameof(axisFont));
        TitleFont = titleFont ?? throw new ArgumentNullException(nameof(titleFont));
        AxisFontSize = ThemeValues.NonNegative(axisFontSize, nameof(axisFontSize));
        TitleFontSize = ThemeValues.NonNegative(titleFontSize, nameof(titleFontSize));
        GridlineColor = gridlineColor;
        ShowDataLabels = showDataLabels;
        ShowLegend = showLegend;
        AxisLineColor = axisLineColor;
        AxisTextColor = axisTextColor;
        GridlineWidth = ThemeValues.NonNegative(gridlineWidth, nameof(gridlineWidth));
        LineWidth = ThemeValues.NonNegative(lineWidth, nameof(lineWidth));
        if (barGapPercent < 0 || barGapPercent > 500) throw new ArgumentOutOfRangeException(nameof(barGapPercent), "The bar gap is 0 to 500 percent.");
        BarGapPercent = barGapPercent;
        Markers = markers;
    }

    /// <summary>Returns a copy with the supplied replacements; omitted values are retained.</summary>
    public OfficeThemeChartStyle With(IReadOnlyList<OfficeColor>? seriesPalette = null, OfficeThemeFont? axisFont = null, OfficeThemeFont? titleFont = null, OfficeLength? axisFontSize = null, OfficeLength? titleFontSize = null, OfficeColor? gridlineColor = null, bool? showDataLabels = null, bool? showLegend = null,
        OfficeColor? axisLineColor = null, OfficeColor? axisTextColor = null, OfficeLength? gridlineWidth = null, OfficeLength? lineWidth = null, int? barGapPercent = null, OfficeChartMarkerShape? markers = null) =>
        new OfficeThemeChartStyle(seriesPalette ?? SeriesPalette, axisFont ?? AxisFont, titleFont ?? TitleFont, axisFontSize ?? AxisFontSize, titleFontSize ?? TitleFontSize, gridlineColor ?? GridlineColor, showDataLabels ?? ShowDataLabels, showLegend ?? ShowLegend,
            axisLineColor ?? AxisLineColor, axisTextColor ?? AxisTextColor, gridlineWidth ?? GridlineWidth, lineWidth ?? LineWidth, barGapPercent ?? BarGapPercent, markers ?? Markers);
}
