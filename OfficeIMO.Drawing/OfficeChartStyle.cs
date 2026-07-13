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
    /// <param name="axisTitleColor">Optional axis title text color.</param>
    /// <param name="titleColor">Chart title color.</param>
    /// <param name="titleFontFamily">Optional chart title font family.</param>
    /// <param name="titleFontSize">Optional chart title font size.</param>
    /// <param name="titleFontStyle">Optional chart title font style.</param>
    /// <param name="plotAreaBackgroundColor">Optional plot area fill.</param>
    /// <param name="plotAreaBorderColor">Optional plot area border color.</param>
    /// <param name="chartBorderWidth">Optional chart border width.</param>
    /// <param name="plotAreaBorderWidth">Optional plot area border width.</param>
    /// <param name="chartBorderDashStyle">Optional chart border dash style.</param>
    /// <param name="plotAreaBorderDashStyle">Optional plot area border dash style.</param>
    /// <param name="showGridLines">Whether cartesian grid lines should be rendered.</param>
    /// <param name="axisLineWidth">Optional axis line width.</param>
    /// <param name="gridLineWidth">Optional grid line width.</param>
    /// <param name="axisLineDashStyle">Optional axis line dash style.</param>
    /// <param name="gridLineDashStyle">Optional grid line dash style.</param>
    /// <param name="legendTextColor">Optional legend text color.</param>
    /// <param name="dataLabelTextColor">Optional data label text color.</param>
    /// <param name="categoryAxisColor">Optional category axis line color.</param>
    /// <param name="valueAxisColor">Optional value axis line color.</param>
    /// <param name="categoryAxisLineWidth">Optional category axis line width.</param>
    /// <param name="valueAxisLineWidth">Optional value axis line width.</param>
    /// <param name="categoryAxisLineDashStyle">Optional category axis line dash style.</param>
    /// <param name="valueAxisLineDashStyle">Optional value axis line dash style.</param>
    /// <param name="categoryGridLineColor">Optional category axis major-gridline color.</param>
    /// <param name="valueGridLineColor">Optional value axis major-gridline color.</param>
    /// <param name="categoryGridLineWidth">Optional category axis major-gridline width.</param>
    /// <param name="valueGridLineWidth">Optional value axis major-gridline width.</param>
    /// <param name="categoryGridLineDashStyle">Optional category axis major-gridline dash style.</param>
    /// <param name="valueGridLineDashStyle">Optional value axis major-gridline dash style.</param>
    /// <param name="showCategoryGridLines">Optional category axis major-gridline visibility.</param>
    /// <param name="showValueGridLines">Optional value axis major-gridline visibility.</param>
    /// <param name="categoryMinorGridLineColor">Optional category axis minor-gridline color.</param>
    /// <param name="valueMinorGridLineColor">Optional value axis minor-gridline color.</param>
    /// <param name="categoryMinorGridLineWidth">Optional category axis minor-gridline width.</param>
    /// <param name="valueMinorGridLineWidth">Optional value axis minor-gridline width.</param>
    /// <param name="categoryMinorGridLineDashStyle">Optional category axis minor-gridline dash style.</param>
    /// <param name="valueMinorGridLineDashStyle">Optional value axis minor-gridline dash style.</param>
    /// <param name="showCategoryMinorGridLines">Optional category axis minor-gridline visibility.</param>
    /// <param name="showValueMinorGridLines">Optional value axis minor-gridline visibility.</param>
    /// <param name="dataLabelFillColor">Optional data label box fill color.</param>
    /// <param name="dataLabelBorderColor">Optional data label box border color.</param>
    /// <param name="dataLabelBorderWidth">Optional data label box border width.</param>
    /// <param name="dataLabelBorderDashStyle">Optional data label box border dash style.</param>
    /// <param name="showBorder">Whether the chart border should be rendered.</param>
    public OfficeChartStyle(
        IEnumerable<OfficeColor>? palette = null,
        string? fontFamily = null,
        OfficeColor? backgroundColor = null,
        OfficeColor? borderColor = null,
        OfficeColor? axisColor = null,
        OfficeColor? gridLineColor = null,
        OfficeColor? textColor = null,
        OfficeColor? mutedTextColor = null,
        OfficeColor? axisTitleColor = null,
        OfficeColor? titleColor = null,
        string? titleFontFamily = null,
        double? titleFontSize = null,
        OfficeFontStyle? titleFontStyle = null,
        OfficeColor? plotAreaBackgroundColor = null,
        OfficeColor? plotAreaBorderColor = null,
        double? chartBorderWidth = null,
        double? plotAreaBorderWidth = null,
        OfficeStrokeDashStyle? chartBorderDashStyle = null,
        OfficeStrokeDashStyle? plotAreaBorderDashStyle = null,
        bool showGridLines = true,
        double? axisLineWidth = null,
        double? gridLineWidth = null,
        OfficeStrokeDashStyle? axisLineDashStyle = null,
        OfficeStrokeDashStyle? gridLineDashStyle = null,
        OfficeColor? legendTextColor = null,
        OfficeColor? dataLabelTextColor = null,
        OfficeColor? categoryAxisColor = null,
        OfficeColor? valueAxisColor = null,
        double? categoryAxisLineWidth = null,
        double? valueAxisLineWidth = null,
        OfficeStrokeDashStyle? categoryAxisLineDashStyle = null,
        OfficeStrokeDashStyle? valueAxisLineDashStyle = null,
        OfficeColor? categoryGridLineColor = null,
        OfficeColor? valueGridLineColor = null,
        double? categoryGridLineWidth = null,
        double? valueGridLineWidth = null,
        OfficeStrokeDashStyle? categoryGridLineDashStyle = null,
        OfficeStrokeDashStyle? valueGridLineDashStyle = null,
        bool? showCategoryGridLines = null,
        bool? showValueGridLines = null,
        OfficeColor? categoryMinorGridLineColor = null,
        OfficeColor? valueMinorGridLineColor = null,
        double? categoryMinorGridLineWidth = null,
        double? valueMinorGridLineWidth = null,
        OfficeStrokeDashStyle? categoryMinorGridLineDashStyle = null,
        OfficeStrokeDashStyle? valueMinorGridLineDashStyle = null,
        bool? showCategoryMinorGridLines = null,
        bool? showValueMinorGridLines = null,
        OfficeColor? dataLabelFillColor = null,
        OfficeColor? dataLabelBorderColor = null,
        double? dataLabelBorderWidth = null,
        OfficeStrokeDashStyle? dataLabelBorderDashStyle = null,
        bool showBorder = true)
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
            axisTitleColor: axisTitleColor,
            titleColor: titleColor,
            titleFontFamily: titleFontFamily,
            titleFontSize: titleFontSize,
            titleFontStyle: titleFontStyle,
            plotAreaBackgroundColor: plotAreaBackgroundColor,
            plotAreaBorderColor: plotAreaBorderColor,
            chartBorderWidth: chartBorderWidth,
            plotAreaBorderWidth: plotAreaBorderWidth,
            chartBorderDashStyle: chartBorderDashStyle,
            plotAreaBorderDashStyle: plotAreaBorderDashStyle,
            showGridLines: showGridLines,
            axisLineWidth: axisLineWidth,
            gridLineWidth: gridLineWidth,
            axisLineDashStyle: axisLineDashStyle,
            gridLineDashStyle: gridLineDashStyle,
            legendTextColor: legendTextColor,
            dataLabelTextColor: dataLabelTextColor,
            categoryAxisColor: categoryAxisColor,
            valueAxisColor: valueAxisColor,
            categoryAxisLineWidth: categoryAxisLineWidth,
            valueAxisLineWidth: valueAxisLineWidth,
            categoryAxisLineDashStyle: categoryAxisLineDashStyle,
            valueAxisLineDashStyle: valueAxisLineDashStyle,
            categoryGridLineColor: categoryGridLineColor,
            valueGridLineColor: valueGridLineColor,
            categoryGridLineWidth: categoryGridLineWidth,
            valueGridLineWidth: valueGridLineWidth,
            categoryGridLineDashStyle: categoryGridLineDashStyle,
            valueGridLineDashStyle: valueGridLineDashStyle,
            showCategoryGridLines: showCategoryGridLines,
            showValueGridLines: showValueGridLines,
            categoryMinorGridLineColor: categoryMinorGridLineColor,
            valueMinorGridLineColor: valueMinorGridLineColor,
            categoryMinorGridLineWidth: categoryMinorGridLineWidth,
            valueMinorGridLineWidth: valueMinorGridLineWidth,
            categoryMinorGridLineDashStyle: categoryMinorGridLineDashStyle,
            valueMinorGridLineDashStyle: valueMinorGridLineDashStyle,
            showCategoryMinorGridLines: showCategoryMinorGridLines,
            showValueMinorGridLines: showValueMinorGridLines,
            dataLabelFillColor: dataLabelFillColor,
            dataLabelBorderColor: dataLabelBorderColor,
            dataLabelBorderWidth: dataLabelBorderWidth,
            dataLabelBorderDashStyle: dataLabelBorderDashStyle,
            showBorder: showBorder) {
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
    /// <param name="axisTitleColor">Optional axis title text color.</param>
    /// <param name="titleColor">Chart title color.</param>
    /// <param name="titleFontFamily">Optional chart title font family.</param>
    /// <param name="titleFontSize">Optional chart title font size.</param>
    /// <param name="titleFontStyle">Optional chart title font style.</param>
    /// <param name="plotAreaBackgroundColor">Optional plot area fill.</param>
    /// <param name="plotAreaBorderColor">Optional plot area border color.</param>
    /// <param name="chartBorderWidth">Optional chart border width.</param>
    /// <param name="plotAreaBorderWidth">Optional plot area border width.</param>
    /// <param name="chartBorderDashStyle">Optional chart border dash style.</param>
    /// <param name="plotAreaBorderDashStyle">Optional plot area border dash style.</param>
    /// <param name="showGridLines">Whether cartesian grid lines should be rendered.</param>
    /// <param name="axisLineWidth">Optional axis line width.</param>
    /// <param name="gridLineWidth">Optional grid line width.</param>
    /// <param name="axisLineDashStyle">Optional axis line dash style.</param>
    /// <param name="gridLineDashStyle">Optional grid line dash style.</param>
    /// <param name="legendTextColor">Optional legend text color.</param>
    /// <param name="dataLabelTextColor">Optional data label text color.</param>
    /// <param name="categoryAxisColor">Optional category axis line color.</param>
    /// <param name="valueAxisColor">Optional value axis line color.</param>
    /// <param name="categoryAxisLineWidth">Optional category axis line width.</param>
    /// <param name="valueAxisLineWidth">Optional value axis line width.</param>
    /// <param name="categoryAxisLineDashStyle">Optional category axis line dash style.</param>
    /// <param name="valueAxisLineDashStyle">Optional value axis line dash style.</param>
    /// <param name="categoryGridLineColor">Optional category axis major-gridline color.</param>
    /// <param name="valueGridLineColor">Optional value axis major-gridline color.</param>
    /// <param name="categoryGridLineWidth">Optional category axis major-gridline width.</param>
    /// <param name="valueGridLineWidth">Optional value axis major-gridline width.</param>
    /// <param name="categoryGridLineDashStyle">Optional category axis major-gridline dash style.</param>
    /// <param name="valueGridLineDashStyle">Optional value axis major-gridline dash style.</param>
    /// <param name="showCategoryGridLines">Optional category axis major-gridline visibility.</param>
    /// <param name="showValueGridLines">Optional value axis major-gridline visibility.</param>
    /// <param name="categoryMinorGridLineColor">Optional category axis minor-gridline color.</param>
    /// <param name="valueMinorGridLineColor">Optional value axis minor-gridline color.</param>
    /// <param name="categoryMinorGridLineWidth">Optional category axis minor-gridline width.</param>
    /// <param name="valueMinorGridLineWidth">Optional value axis minor-gridline width.</param>
    /// <param name="categoryMinorGridLineDashStyle">Optional category axis minor-gridline dash style.</param>
    /// <param name="valueMinorGridLineDashStyle">Optional value axis minor-gridline dash style.</param>
    /// <param name="showCategoryMinorGridLines">Optional category axis minor-gridline visibility.</param>
    /// <param name="showValueMinorGridLines">Optional value axis minor-gridline visibility.</param>
    /// <param name="dataLabelFillColor">Optional data label box fill color.</param>
    /// <param name="dataLabelBorderColor">Optional data label box border color.</param>
    /// <param name="dataLabelBorderWidth">Optional data label box border width.</param>
    /// <param name="dataLabelBorderDashStyle">Optional data label box border dash style.</param>
    /// <param name="showBorder">Whether the chart border should be rendered.</param>
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
        OfficeColor? axisTitleColor = null,
        OfficeColor? titleColor = null,
        string? titleFontFamily = null,
        double? titleFontSize = null,
        OfficeFontStyle? titleFontStyle = null,
        OfficeColor? plotAreaBackgroundColor = null,
        OfficeColor? plotAreaBorderColor = null,
        double? chartBorderWidth = null,
        double? plotAreaBorderWidth = null,
        OfficeStrokeDashStyle? chartBorderDashStyle = null,
        OfficeStrokeDashStyle? plotAreaBorderDashStyle = null,
        bool showGridLines = true,
        double? axisLineWidth = null,
        double? gridLineWidth = null,
        OfficeStrokeDashStyle? axisLineDashStyle = null,
        OfficeStrokeDashStyle? gridLineDashStyle = null,
        OfficeColor? legendTextColor = null,
        OfficeColor? dataLabelTextColor = null,
        OfficeColor? categoryAxisColor = null,
        OfficeColor? valueAxisColor = null,
        double? categoryAxisLineWidth = null,
        double? valueAxisLineWidth = null,
        OfficeStrokeDashStyle? categoryAxisLineDashStyle = null,
        OfficeStrokeDashStyle? valueAxisLineDashStyle = null,
        OfficeColor? categoryGridLineColor = null,
        OfficeColor? valueGridLineColor = null,
        double? categoryGridLineWidth = null,
        double? valueGridLineWidth = null,
        OfficeStrokeDashStyle? categoryGridLineDashStyle = null,
        OfficeStrokeDashStyle? valueGridLineDashStyle = null,
        bool? showCategoryGridLines = null,
        bool? showValueGridLines = null,
        OfficeColor? categoryMinorGridLineColor = null,
        OfficeColor? valueMinorGridLineColor = null,
        double? categoryMinorGridLineWidth = null,
        double? valueMinorGridLineWidth = null,
        OfficeStrokeDashStyle? categoryMinorGridLineDashStyle = null,
        OfficeStrokeDashStyle? valueMinorGridLineDashStyle = null,
        bool? showCategoryMinorGridLines = null,
        bool? showValueMinorGridLines = null,
        OfficeColor? dataLabelFillColor = null,
        OfficeColor? dataLabelBorderColor = null,
        double? dataLabelBorderWidth = null,
        OfficeStrokeDashStyle? dataLabelBorderDashStyle = null,
        bool showBorder = true) {
        if (chartBorderWidth is <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(chartBorderWidth), "Chart border width must be greater than zero.");
        }
        if (plotAreaBorderWidth is <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(plotAreaBorderWidth), "Plot area border width must be greater than zero.");
        }
        if (axisLineWidth is <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(axisLineWidth), "Axis line width must be greater than zero.");
        }
        if (categoryAxisLineWidth is <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(categoryAxisLineWidth), "Category axis line width must be greater than zero.");
        }
        if (valueAxisLineWidth is <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(valueAxisLineWidth), "Value axis line width must be greater than zero.");
        }
        if (gridLineWidth is <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(gridLineWidth), "Grid line width must be greater than zero.");
        }
        if (categoryGridLineWidth is <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(categoryGridLineWidth), "Category gridline width must be greater than zero.");
        }
        if (valueGridLineWidth is <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(valueGridLineWidth), "Value gridline width must be greater than zero.");
        }
        if (categoryMinorGridLineWidth is <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(categoryMinorGridLineWidth), "Category minor gridline width must be greater than zero.");
        }
        if (valueMinorGridLineWidth is <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(valueMinorGridLineWidth), "Value minor gridline width must be greater than zero.");
        }
        if (titleFontSize is <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(titleFontSize), "Chart title font size must be greater than zero.");
        }
        if (dataLabelBorderWidth is <= 0D) {
            throw new ArgumentOutOfRangeException(nameof(dataLabelBorderWidth), "Data label border width must be greater than zero.");
        }

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
        AxisTitleColor = axisTitleColor;
        TitleColor = titleColor ?? OfficeColor.FromRgb(31, 78, 121);
        TitleFontFamily = string.IsNullOrWhiteSpace(titleFontFamily) ? null : titleFontFamily!;
        TitleFontSize = titleFontSize;
        TitleFontStyle = titleFontStyle;
        PlotAreaBackgroundColor = plotAreaBackgroundColor;
        PlotAreaBorderColor = plotAreaBorderColor;
        ChartBorderWidth = chartBorderWidth;
        PlotAreaBorderWidth = plotAreaBorderWidth;
        ChartBorderDashStyle = chartBorderDashStyle;
        PlotAreaBorderDashStyle = plotAreaBorderDashStyle;
        ShowGridLines = showGridLines;
        AxisLineWidth = axisLineWidth;
        CategoryAxisColor = categoryAxisColor;
        ValueAxisColor = valueAxisColor;
        CategoryAxisLineWidth = categoryAxisLineWidth;
        ValueAxisLineWidth = valueAxisLineWidth;
        CategoryAxisLineDashStyle = categoryAxisLineDashStyle;
        ValueAxisLineDashStyle = valueAxisLineDashStyle;
        GridLineWidth = gridLineWidth;
        AxisLineDashStyle = axisLineDashStyle;
        GridLineDashStyle = gridLineDashStyle;
        CategoryGridLineColor = categoryGridLineColor;
        ValueGridLineColor = valueGridLineColor;
        CategoryGridLineWidth = categoryGridLineWidth;
        ValueGridLineWidth = valueGridLineWidth;
        CategoryGridLineDashStyle = categoryGridLineDashStyle;
        ValueGridLineDashStyle = valueGridLineDashStyle;
        ShowCategoryGridLines = showCategoryGridLines;
        ShowValueGridLines = showValueGridLines;
        CategoryMinorGridLineColor = categoryMinorGridLineColor;
        ValueMinorGridLineColor = valueMinorGridLineColor;
        CategoryMinorGridLineWidth = categoryMinorGridLineWidth;
        ValueMinorGridLineWidth = valueMinorGridLineWidth;
        CategoryMinorGridLineDashStyle = categoryMinorGridLineDashStyle;
        ValueMinorGridLineDashStyle = valueMinorGridLineDashStyle;
        ShowCategoryMinorGridLines = showCategoryMinorGridLines;
        ShowValueMinorGridLines = showValueMinorGridLines;
        ShowBackground = showBackground;
        ShowBorder = showBorder;
        LegendTextColor = legendTextColor;
        DataLabelTextColor = dataLabelTextColor;
        DataLabelFillColor = dataLabelFillColor;
        DataLabelBorderColor = dataLabelBorderColor;
        DataLabelBorderWidth = dataLabelBorderWidth;
        DataLabelBorderDashStyle = dataLabelBorderDashStyle;
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

    /// <summary>Optional category axis line color override.</summary>
    public OfficeColor? CategoryAxisColor { get; }

    /// <summary>Optional value axis line color override.</summary>
    public OfficeColor? ValueAxisColor { get; }

    /// <summary>Grid line color.</summary>
    public OfficeColor GridLineColor { get; }

    /// <summary>Optional category axis major-gridline color override.</summary>
    public OfficeColor? CategoryGridLineColor { get; }

    /// <summary>Optional value axis major-gridline color override.</summary>
    public OfficeColor? ValueGridLineColor { get; }

    /// <summary>Optional category axis minor-gridline color override.</summary>
    public OfficeColor? CategoryMinorGridLineColor { get; }

    /// <summary>Optional value axis minor-gridline color override.</summary>
    public OfficeColor? ValueMinorGridLineColor { get; }

    /// <summary>Primary chart text color.</summary>
    public OfficeColor TextColor { get; }

    /// <summary>Optional legend text color.</summary>
    public OfficeColor? LegendTextColor { get; }

    /// <summary>Optional data label text color.</summary>
    public OfficeColor? DataLabelTextColor { get; }

    /// <summary>Optional data label box fill color.</summary>
    public OfficeColor? DataLabelFillColor { get; }

    /// <summary>Optional data label box border color.</summary>
    public OfficeColor? DataLabelBorderColor { get; }

    /// <summary>Optional data label box border width.</summary>
    public double? DataLabelBorderWidth { get; }

    /// <summary>Optional data label box border dash style.</summary>
    public OfficeStrokeDashStyle? DataLabelBorderDashStyle { get; }

    /// <summary>Secondary chart text color for axis and category labels.</summary>
    public OfficeColor MutedTextColor { get; }

    /// <summary>Optional axis title text color.</summary>
    public OfficeColor? AxisTitleColor { get; }

    /// <summary>Chart title color.</summary>
    public OfficeColor TitleColor { get; }

    /// <summary>Optional chart title font family.</summary>
    public string? TitleFontFamily { get; }

    /// <summary>Optional chart title font size.</summary>
    public double? TitleFontSize { get; }

    /// <summary>Optional chart title font style.</summary>
    public OfficeFontStyle? TitleFontStyle { get; }

    /// <summary>Optional plot area fill.</summary>
    public OfficeColor? PlotAreaBackgroundColor { get; }

    /// <summary>Optional plot area border color.</summary>
    public OfficeColor? PlotAreaBorderColor { get; }

    /// <summary>Optional chart border width.</summary>
    public double? ChartBorderWidth { get; }

    /// <summary>Optional plot area border width.</summary>
    public double? PlotAreaBorderWidth { get; }

    /// <summary>Optional chart border dash style.</summary>
    public OfficeStrokeDashStyle? ChartBorderDashStyle { get; }

    /// <summary>Optional plot area border dash style.</summary>
    public OfficeStrokeDashStyle? PlotAreaBorderDashStyle { get; }

    /// <summary>Whether cartesian grid lines should be rendered.</summary>
    public bool ShowGridLines { get; }

    /// <summary>Optional category axis major-gridline visibility override.</summary>
    public bool? ShowCategoryGridLines { get; }

    /// <summary>Optional value axis major-gridline visibility override.</summary>
    public bool? ShowValueGridLines { get; }

    /// <summary>Optional category axis minor-gridline visibility override.</summary>
    public bool? ShowCategoryMinorGridLines { get; }

    /// <summary>Optional value axis minor-gridline visibility override.</summary>
    public bool? ShowValueMinorGridLines { get; }

    /// <summary>Optional axis line width.</summary>
    public double? AxisLineWidth { get; }

    /// <summary>Optional category axis line width override.</summary>
    public double? CategoryAxisLineWidth { get; }

    /// <summary>Optional value axis line width override.</summary>
    public double? ValueAxisLineWidth { get; }

    /// <summary>Optional grid line width.</summary>
    public double? GridLineWidth { get; }

    /// <summary>Optional category axis major-gridline width override.</summary>
    public double? CategoryGridLineWidth { get; }

    /// <summary>Optional value axis major-gridline width override.</summary>
    public double? ValueGridLineWidth { get; }

    /// <summary>Optional category axis minor-gridline width override.</summary>
    public double? CategoryMinorGridLineWidth { get; }

    /// <summary>Optional value axis minor-gridline width override.</summary>
    public double? ValueMinorGridLineWidth { get; }

    /// <summary>Optional axis line dash style.</summary>
    public OfficeStrokeDashStyle? AxisLineDashStyle { get; }

    /// <summary>Optional category axis line dash style override.</summary>
    public OfficeStrokeDashStyle? CategoryAxisLineDashStyle { get; }

    /// <summary>Optional value axis line dash style override.</summary>
    public OfficeStrokeDashStyle? ValueAxisLineDashStyle { get; }

    /// <summary>Optional grid line dash style.</summary>
    public OfficeStrokeDashStyle? GridLineDashStyle { get; }

    /// <summary>Optional category axis major-gridline dash style override.</summary>
    public OfficeStrokeDashStyle? CategoryGridLineDashStyle { get; }

    /// <summary>Optional value axis major-gridline dash style override.</summary>
    public OfficeStrokeDashStyle? ValueGridLineDashStyle { get; }

    /// <summary>Optional category axis minor-gridline dash style override.</summary>
    public OfficeStrokeDashStyle? CategoryMinorGridLineDashStyle { get; }

    /// <summary>Optional value axis minor-gridline dash style override.</summary>
    public OfficeStrokeDashStyle? ValueMinorGridLineDashStyle { get; }

    /// <summary>Whether the chart background fill should be rendered.</summary>
    public bool ShowBackground { get; }

    /// <summary>Whether the chart border should be rendered.</summary>
    public bool ShowBorder { get; }

    /// <summary>Gets a palette color for the zero-based series or slice index.</summary>
    public OfficeColor GetSeriesColor(int index) {
        int count = Palette.Count;
        int normalized = count == 0 ? 0 : (int)(Math.Abs((long)index) % count);
        return count == 0 ? DefaultPaletteValues[0] : Palette[normalized];
    }
}
