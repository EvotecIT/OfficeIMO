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

    /// <summary>Creates immutable chart authoring tokens, separate from legacy renderer overrides.</summary>
    public OfficeThemeChartStyle(IReadOnlyList<OfficeColor> seriesPalette, OfficeThemeFont axisFont, OfficeThemeFont titleFont, OfficeLength axisFontSize, OfficeLength titleFontSize, OfficeColor gridlineColor, bool showDataLabels, bool showLegend) {
        SeriesPalette = ThemeValues.Colors(seriesPalette, nameof(seriesPalette));
        AxisFont = axisFont ?? throw new ArgumentNullException(nameof(axisFont));
        TitleFont = titleFont ?? throw new ArgumentNullException(nameof(titleFont));
        AxisFontSize = ThemeValues.NonNegative(axisFontSize, nameof(axisFontSize));
        TitleFontSize = ThemeValues.NonNegative(titleFontSize, nameof(titleFontSize));
        GridlineColor = gridlineColor;
        ShowDataLabels = showDataLabels;
        ShowLegend = showLegend;
    }

    /// <summary>Returns a copy with the supplied replacements; omitted values are retained.</summary>
    public OfficeThemeChartStyle With(IReadOnlyList<OfficeColor>? seriesPalette = null, OfficeThemeFont? axisFont = null, OfficeThemeFont? titleFont = null, OfficeLength? axisFontSize = null, OfficeLength? titleFontSize = null, OfficeColor? gridlineColor = null, bool? showDataLabels = null, bool? showLegend = null) =>
        new OfficeThemeChartStyle(seriesPalette ?? SeriesPalette, axisFont ?? AxisFont, titleFont ?? TitleFont, axisFontSize ?? AxisFontSize, titleFontSize ?? TitleFontSize, gridlineColor ?? GridlineColor, showDataLabels ?? ShowDataLabels, showLegend ?? ShowLegend);
}
