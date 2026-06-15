using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>
/// Reusable chart style metadata shared by OfficeIMO chart renderers and format exporters.
/// </summary>
public sealed class OfficeChartStyle {
    private static readonly OfficeColor[] DefaultPaletteValues = new[] {
        OfficeColor.FromRgb(31, 78, 121),
        OfficeColor.FromRgb(47, 111, 62),
        OfficeColor.FromRgb(184, 90, 35),
        OfficeColor.FromRgb(112, 48, 160),
        OfficeColor.FromRgb(37, 99, 235),
        OfficeColor.FromRgb(120, 113, 108)
    };

    private static readonly OfficeChartStyle DefaultStyle = new OfficeChartStyle();

    /// <summary>
    /// Creates chart style metadata.
    /// </summary>
    /// <param name="palette">Optional series or slice palette. Empty palettes fall back to the default premium palette.</param>
    /// <param name="fontFamily">Optional chart text font family. Blank values fall back to Aptos.</param>
    /// <param name="backgroundColor">Chart background fill.</param>
    /// <param name="borderColor">Chart border color.</param>
    /// <param name="axisColor">Axis line color.</param>
    /// <param name="gridLineColor">Grid line color.</param>
    /// <param name="textColor">Primary chart text color.</param>
    /// <param name="mutedTextColor">Secondary chart text color for axis and category labels.</param>
    /// <param name="titleColor">Chart title color.</param>
    /// <param name="plotAreaBackgroundColor">Optional plot area fill.</param>
    /// <param name="plotAreaBorderColor">Optional plot area border color.</param>
    /// <param name="showGridLines">Whether cartesian grid lines should be rendered.</param>
    public OfficeChartStyle(
        IEnumerable<OfficeColor>? palette = null,
        string? fontFamily = null,
        OfficeColor? backgroundColor = null,
        OfficeColor? borderColor = null,
        OfficeColor? axisColor = null,
        OfficeColor? gridLineColor = null,
        OfficeColor? textColor = null,
        OfficeColor? mutedTextColor = null,
        OfficeColor? titleColor = null,
        OfficeColor? plotAreaBackgroundColor = null,
        OfficeColor? plotAreaBorderColor = null,
        bool showGridLines = true)
        : this(
            showBackground: true,
            palette: palette,
            fontFamily: fontFamily,
            backgroundColor: backgroundColor,
            borderColor: borderColor,
            axisColor: axisColor,
            gridLineColor: gridLineColor,
            textColor: textColor,
            mutedTextColor: mutedTextColor,
            titleColor: titleColor,
            plotAreaBackgroundColor: plotAreaBackgroundColor,
            plotAreaBorderColor: plotAreaBorderColor,
            showGridLines: showGridLines) {
    }

    /// <summary>
    /// Creates chart style metadata with explicit background visibility.
    /// </summary>
    /// <param name="showBackground">Whether the chart background fill should be rendered.</param>
    /// <param name="palette">Optional series or slice palette. Empty palettes fall back to the default premium palette.</param>
    /// <param name="fontFamily">Optional chart text font family. Blank values fall back to Aptos.</param>
    /// <param name="backgroundColor">Chart background fill.</param>
    /// <param name="borderColor">Chart border color.</param>
    /// <param name="axisColor">Axis line color.</param>
    /// <param name="gridLineColor">Grid line color.</param>
    /// <param name="textColor">Primary chart text color.</param>
    /// <param name="mutedTextColor">Secondary chart text color for axis and category labels.</param>
    /// <param name="titleColor">Chart title color.</param>
    /// <param name="plotAreaBackgroundColor">Optional plot area fill.</param>
    /// <param name="plotAreaBorderColor">Optional plot area border color.</param>
    /// <param name="showGridLines">Whether cartesian grid lines should be rendered.</param>
    public OfficeChartStyle(
        bool showBackground,
        IEnumerable<OfficeColor>? palette = null,
        string? fontFamily = null,
        OfficeColor? backgroundColor = null,
        OfficeColor? borderColor = null,
        OfficeColor? axisColor = null,
        OfficeColor? gridLineColor = null,
        OfficeColor? textColor = null,
        OfficeColor? mutedTextColor = null,
        OfficeColor? titleColor = null,
        OfficeColor? plotAreaBackgroundColor = null,
        OfficeColor? plotAreaBorderColor = null,
        bool showGridLines = true) {
        var colors = palette == null ? new List<OfficeColor>() : new List<OfficeColor>(palette);
        if (colors.Count == 0) {
            colors.AddRange(DefaultPaletteValues);
        }

        Palette = new ReadOnlyCollection<OfficeColor>(colors);
        FontFamily = string.IsNullOrWhiteSpace(fontFamily) ? "Aptos" : fontFamily!;
        BackgroundColor = backgroundColor ?? OfficeColor.FromRgb(250, 252, 255);
        BorderColor = borderColor ?? OfficeColor.FromRgb(183, 194, 207);
        AxisColor = axisColor ?? OfficeColor.FromRgb(80, 90, 100);
        GridLineColor = gridLineColor ?? OfficeColor.FromRgb(226, 232, 240);
        TextColor = textColor ?? OfficeColor.FromRgb(51, 65, 85);
        MutedTextColor = mutedTextColor ?? OfficeColor.FromRgb(100, 116, 139);
        TitleColor = titleColor ?? OfficeColor.FromRgb(31, 78, 121);
        PlotAreaBackgroundColor = plotAreaBackgroundColor;
        PlotAreaBorderColor = plotAreaBorderColor;
        ShowGridLines = showGridLines;
        ShowBackground = showBackground;
        ShowBorder = true;
    }

    /// <summary>Default premium OfficeIMO chart style.</summary>
    public static OfficeChartStyle Default => DefaultStyle;

    /// <summary>Series and slice palette.</summary>
    public IReadOnlyList<OfficeColor> Palette { get; }

    /// <summary>Chart text font family.</summary>
    public string FontFamily { get; }

    /// <summary>Chart background fill.</summary>
    public OfficeColor BackgroundColor { get; }

    /// <summary>Chart border color.</summary>
    public OfficeColor BorderColor { get; }

    /// <summary>Axis line color.</summary>
    public OfficeColor AxisColor { get; }

    /// <summary>Grid line color.</summary>
    public OfficeColor GridLineColor { get; }

    /// <summary>Primary chart text color.</summary>
    public OfficeColor TextColor { get; }

    /// <summary>Secondary chart text color for axis and category labels.</summary>
    public OfficeColor MutedTextColor { get; }

    /// <summary>Chart title color.</summary>
    public OfficeColor TitleColor { get; }

    /// <summary>Optional plot area fill.</summary>
    public OfficeColor? PlotAreaBackgroundColor { get; }

    /// <summary>Optional plot area border color.</summary>
    public OfficeColor? PlotAreaBorderColor { get; }

    /// <summary>Whether cartesian grid lines should be rendered.</summary>
    public bool ShowGridLines { get; }

    /// <summary>Whether the chart background fill should be rendered.</summary>
    public bool ShowBackground { get; }

    /// <summary>Whether the chart border should be rendered.</summary>
    public bool ShowBorder { get; set; }

    /// <summary>Gets a palette color for the zero-based series or slice index.</summary>
    public OfficeColor GetSeriesColor(int index) {
        int count = Palette.Count;
        int normalized = count == 0 ? 0 : (int)(Math.Abs((long)index) % count);
        return count == 0 ? DefaultPaletteValues[0] : Palette[normalized];
    }
}
