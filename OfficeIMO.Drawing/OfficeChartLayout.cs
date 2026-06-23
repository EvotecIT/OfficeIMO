using System;

using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Reusable chart layout metadata shared by OfficeIMO chart renderers and format exporters.
/// </summary>
public sealed class OfficeChartLayout {
    internal const int MaxNumberFormatLength = 1024;

    private static readonly OfficeChartLayout DefaultLayout = new OfficeChartLayout();

    /// <summary>
    /// Creates chart layout metadata.
    /// </summary>
    /// <param name="seriesLegendWidthRatio">Maximum chart-width ratio reserved for series legends.</param>
    /// <param name="categoryLegendWidthRatio">Maximum chart-width ratio reserved for category legends such as pie slices.</param>
    /// <param name="legendRowHeight">Legend row height.</param>
    /// <param name="legendSwatchSize">Legend color swatch size.</param>
    /// <param name="legendTextGap">Gap between a legend swatch and its label.</param>
    /// <param name="legendFontSize">Legend label font size.</param>
    /// <param name="legendFontFamily">Optional legend label font family.</param>
    /// <param name="axisLabelFontSize">Axis label font size.</param>
    /// <param name="axisTextFontFamily">Optional axis label/title font family.</param>
    /// <param name="categoryAxisLabelWidth">Maximum category-axis label width.</param>
    /// <param name="radarCategoryLabelWidth">Maximum radar category label width.</param>
    /// <param name="maximumCategoryAxisLabels">Maximum number of category-axis labels to render on cartesian charts.</param>
    /// <param name="maximumHorizontalCategoryAxisLabels">Maximum number of category-axis labels to render on horizontal bar charts.</param>
    /// <param name="maximumRadarCategoryLabels">Maximum number of category labels to render on radar charts.</param>
    /// <param name="preventLabelOverlap">Whether axis/category label stride should increase automatically to avoid obvious label overlap.</param>
    /// <param name="showLegend">Whether series or category legends should be rendered.</param>
    /// <param name="legendPosition">Preferred legend placement when legends are rendered.</param>
    /// <param name="showDataLabels">Whether point data labels should be rendered when supported by a chart family.</param>
    /// <param name="showDataLabelValues">Whether point data labels should include values.</param>
    /// <param name="showDataLabelPercentages">Whether point data labels should include percentages.</param>
    /// <param name="showDataLabelCategoryNames">Whether point data labels should include category names.</param>
    /// <param name="showDataLabelSeriesNames">Whether point data labels should include series names.</param>
    /// <param name="dataLabelSeparator">Separator used between enabled data label parts.</param>
    /// <param name="dataLabelFontSize">Data label font size.</param>
    /// <param name="dataLabelFontFamily">Optional data label font family.</param>
    /// <param name="dataLabelPosition">Preferred data label position when labels are rendered.</param>
    /// <param name="dataLabelNumberFormat">Optional numeric format for data-label values.</param>
    /// <param name="showMarkers">Whether point markers should be rendered for marker-capable chart families.</param>
    /// <param name="axisNumberFormat">Optional numeric format for value-axis labels.</param>
    /// <param name="categoryAxisTitle">Optional category or horizontal axis title.</param>
    /// <param name="valueAxisTitle">Optional value or vertical axis title.</param>
    /// <param name="horizontalAxisNumberFormat">Optional numeric format for horizontal value-axis labels.</param>
    /// <param name="verticalAxisNumberFormat">Optional numeric format for vertical value-axis labels.</param>
    /// <param name="connectScatterPoints">Whether scatter points should be connected by series lines.</param>
    /// <param name="fillRadarSeries">Whether radar series polygons should be filled.</param>
    /// <param name="showCategoryAxis">Whether the category or horizontal axis should be rendered.</param>
    /// <param name="showValueAxis">Whether the value or vertical axis should be rendered.</param>
    /// <param name="showCategoryAxisLine">Whether the category or horizontal axis line should be rendered.</param>
    /// <param name="showValueAxisLine">Whether the value or vertical axis line should be rendered.</param>
    /// <param name="showCategoryAxisLabels">Whether category or horizontal tick labels should be rendered.</param>
    /// <param name="showValueAxisLabels">Whether value or vertical tick labels should be rendered.</param>
    /// <param name="horizontalAxisTickLabelPosition">Preferred tick-label side for the physical horizontal axis.</param>
    /// <param name="verticalAxisTickLabelPosition">Preferred tick-label side for the physical vertical axis.</param>
    /// <param name="horizontalAxisCrossingPosition">Physical crossing side for the horizontal axis.</param>
    /// <param name="verticalAxisCrossingPosition">Physical crossing side for the vertical axis.</param>
    /// <param name="reverseCategoryAxis">Whether category-axis slots should be rendered in reverse order.</param>
    /// <param name="overlayTitle">Whether the title should overlay the plot instead of reserving layout space.</param>
    /// <param name="titleTopPadding">Top padding before the chart title inside the chart canvas.</param>
    /// <param name="legendFontStyle">Optional legend label font style.</param>
    /// <param name="axisTextFontStyle">Optional axis label font style and axis title fallback style.</param>
    /// <param name="dataLabelFontStyle">Optional data label font style.</param>
    /// <param name="axisTitleFontSize">Optional axis title font size.</param>
    /// <param name="axisTitleFontFamily">Optional axis title font family.</param>
    /// <param name="axisTitleFontStyle">Optional axis title font style.</param>
    /// <param name="horizontalAxisDisplayUnitDivisor">Optional divisor applied to horizontal value-axis labels.</param>
    /// <param name="horizontalAxisDisplayUnitLabel">Optional display-unit label shown for horizontal value axes.</param>
    /// <param name="verticalAxisDisplayUnitDivisor">Optional divisor applied to vertical value-axis labels.</param>
    /// <param name="verticalAxisDisplayUnitLabel">Optional display-unit label shown for vertical value axes.</param>
    /// <param name="horizontalAxisMinimum">Optional minimum for horizontal value axes.</param>
    /// <param name="horizontalAxisMaximum">Optional maximum for horizontal value axes.</param>
    /// <param name="horizontalAxisMajorUnit">Optional major tick/grid unit for horizontal value axes.</param>
    /// <param name="horizontalAxisMinorUnit">Optional minor tick/grid unit for horizontal value axes.</param>
    /// <param name="verticalAxisMinimum">Optional minimum for vertical value axes.</param>
    /// <param name="verticalAxisMaximum">Optional maximum for vertical value axes.</param>
    /// <param name="verticalAxisMajorUnit">Optional major tick/grid unit for vertical value axes.</param>
    /// <param name="verticalAxisMinorUnit">Optional minor tick/grid unit for vertical value axes.</param>
    /// <param name="horizontalAxisMajorTickMark">Optional major tick mark placement for the horizontal axis.</param>
    /// <param name="verticalAxisMajorTickMark">Optional major tick mark placement for the vertical axis.</param>
    /// <param name="horizontalAxisMinorTickMark">Optional minor tick mark placement for the horizontal axis.</param>
    /// <param name="verticalAxisMinorTickMark">Optional minor tick mark placement for the vertical axis.</param>
    public OfficeChartLayout(
        double? seriesLegendWidthRatio = null,
        double? categoryLegendWidthRatio = null,
        double? legendRowHeight = null,
        double? legendSwatchSize = null,
        double? legendTextGap = null,
        double? legendFontSize = null,
        string? legendFontFamily = null,
        double? axisLabelFontSize = null,
        string? axisTextFontFamily = null,
        double? categoryAxisLabelWidth = null,
        double? radarCategoryLabelWidth = null,
        int? maximumCategoryAxisLabels = null,
        int? maximumHorizontalCategoryAxisLabels = null,
        int? maximumRadarCategoryLabels = null,
        bool preventLabelOverlap = true,
        bool showLegend = true,
        OfficeChartLegendPosition legendPosition = OfficeChartLegendPosition.Right,
        bool showDataLabels = false,
        bool showDataLabelValues = false,
        bool showDataLabelPercentages = false,
        bool showDataLabelCategoryNames = false,
        bool showDataLabelSeriesNames = false,
        string? dataLabelSeparator = null,
        double? dataLabelFontSize = null,
        string? dataLabelFontFamily = null,
        OfficeChartDataLabelPosition dataLabelPosition = OfficeChartDataLabelPosition.Auto,
        string? dataLabelNumberFormat = null,
        bool showMarkers = true,
        string? axisNumberFormat = null,
        string? categoryAxisTitle = null,
        string? valueAxisTitle = null,
        string? horizontalAxisNumberFormat = null,
        string? verticalAxisNumberFormat = null,
        bool connectScatterPoints = true,
        bool fillRadarSeries = true,
        bool showCategoryAxis = true,
        bool showValueAxis = true,
        bool showCategoryAxisLine = true,
        bool showValueAxisLine = true,
        bool showCategoryAxisLabels = true,
        bool showValueAxisLabels = true,
        OfficeChartAxisTickLabelPosition horizontalAxisTickLabelPosition = OfficeChartAxisTickLabelPosition.NextTo,
        OfficeChartAxisTickLabelPosition verticalAxisTickLabelPosition = OfficeChartAxisTickLabelPosition.NextTo,
        OfficeChartAxisCrossingPosition horizontalAxisCrossingPosition = OfficeChartAxisCrossingPosition.AutoZero,
        OfficeChartAxisCrossingPosition verticalAxisCrossingPosition = OfficeChartAxisCrossingPosition.AutoZero,
        bool reverseCategoryAxis = false,
        bool overlayTitle = false,
        double? titleTopPadding = null,
        OfficeFontStyle? legendFontStyle = null,
        OfficeFontStyle? axisTextFontStyle = null,
        OfficeFontStyle? dataLabelFontStyle = null,
        double? axisTitleFontSize = null,
        string? axisTitleFontFamily = null,
        OfficeFontStyle? axisTitleFontStyle = null,
        double? horizontalAxisDisplayUnitDivisor = null,
        string? horizontalAxisDisplayUnitLabel = null,
        double? verticalAxisDisplayUnitDivisor = null,
        string? verticalAxisDisplayUnitLabel = null,
        double? horizontalAxisMinimum = null,
        double? horizontalAxisMaximum = null,
        double? horizontalAxisMajorUnit = null,
        double? horizontalAxisMinorUnit = null,
        double? verticalAxisMinimum = null,
        double? verticalAxisMaximum = null,
        double? verticalAxisMajorUnit = null,
        double? verticalAxisMinorUnit = null,
        OfficeChartAxisTickMark horizontalAxisMajorTickMark = OfficeChartAxisTickMark.None,
        OfficeChartAxisTickMark verticalAxisMajorTickMark = OfficeChartAxisTickMark.None,
        OfficeChartAxisTickMark horizontalAxisMinorTickMark = OfficeChartAxisTickMark.None,
        OfficeChartAxisTickMark verticalAxisMinorTickMark = OfficeChartAxisTickMark.None)
        : this(
            overlayLegend: false,
            seriesLegendWidthRatio: seriesLegendWidthRatio,
            categoryLegendWidthRatio: categoryLegendWidthRatio,
            legendRowHeight: legendRowHeight,
            legendSwatchSize: legendSwatchSize,
            legendTextGap: legendTextGap,
            legendFontSize: legendFontSize,
            legendFontFamily: legendFontFamily,
            axisLabelFontSize: axisLabelFontSize,
            axisTextFontFamily: axisTextFontFamily,
            categoryAxisLabelWidth: categoryAxisLabelWidth,
            radarCategoryLabelWidth: radarCategoryLabelWidth,
            maximumCategoryAxisLabels: maximumCategoryAxisLabels,
            maximumHorizontalCategoryAxisLabels: maximumHorizontalCategoryAxisLabels,
            maximumRadarCategoryLabels: maximumRadarCategoryLabels,
            preventLabelOverlap: preventLabelOverlap,
            showLegend: showLegend,
            legendPosition: legendPosition,
            showDataLabels: showDataLabels,
            showDataLabelValues: showDataLabelValues,
            showDataLabelPercentages: showDataLabelPercentages,
            showDataLabelCategoryNames: showDataLabelCategoryNames,
            showDataLabelSeriesNames: showDataLabelSeriesNames,
            dataLabelSeparator: dataLabelSeparator,
            dataLabelFontSize: dataLabelFontSize,
            dataLabelFontFamily: dataLabelFontFamily,
            dataLabelPosition: dataLabelPosition,
            dataLabelNumberFormat: dataLabelNumberFormat,
            showMarkers: showMarkers,
            axisNumberFormat: axisNumberFormat,
            categoryAxisTitle: categoryAxisTitle,
            valueAxisTitle: valueAxisTitle,
            horizontalAxisNumberFormat: horizontalAxisNumberFormat,
            verticalAxisNumberFormat: verticalAxisNumberFormat,
            connectScatterPoints: connectScatterPoints,
            fillRadarSeries: fillRadarSeries,
            showCategoryAxis: showCategoryAxis,
            showValueAxis: showValueAxis,
            showCategoryAxisLine: showCategoryAxisLine,
            showValueAxisLine: showValueAxisLine,
            showCategoryAxisLabels: showCategoryAxisLabels,
            showValueAxisLabels: showValueAxisLabels,
            horizontalAxisTickLabelPosition: horizontalAxisTickLabelPosition,
            verticalAxisTickLabelPosition: verticalAxisTickLabelPosition,
            horizontalAxisCrossingPosition: horizontalAxisCrossingPosition,
            verticalAxisCrossingPosition: verticalAxisCrossingPosition,
            reverseCategoryAxis: reverseCategoryAxis,
            overlayTitle: overlayTitle,
            titleTopPadding: titleTopPadding,
            legendFontStyle: legendFontStyle,
            axisTextFontStyle: axisTextFontStyle,
            dataLabelFontStyle: dataLabelFontStyle,
            axisTitleFontSize: axisTitleFontSize,
            axisTitleFontFamily: axisTitleFontFamily,
            axisTitleFontStyle: axisTitleFontStyle,
            horizontalAxisDisplayUnitDivisor: horizontalAxisDisplayUnitDivisor,
            horizontalAxisDisplayUnitLabel: horizontalAxisDisplayUnitLabel,
            verticalAxisDisplayUnitDivisor: verticalAxisDisplayUnitDivisor,
            verticalAxisDisplayUnitLabel: verticalAxisDisplayUnitLabel,
            horizontalAxisMinimum: horizontalAxisMinimum,
            horizontalAxisMaximum: horizontalAxisMaximum,
            horizontalAxisMajorUnit: horizontalAxisMajorUnit,
            horizontalAxisMinorUnit: horizontalAxisMinorUnit,
            verticalAxisMinimum: verticalAxisMinimum,
            verticalAxisMaximum: verticalAxisMaximum,
            verticalAxisMajorUnit: verticalAxisMajorUnit,
            verticalAxisMinorUnit: verticalAxisMinorUnit,
            horizontalAxisMajorTickMark: horizontalAxisMajorTickMark,
            verticalAxisMajorTickMark: verticalAxisMajorTickMark,
            horizontalAxisMinorTickMark: horizontalAxisMinorTickMark,
            verticalAxisMinorTickMark: verticalAxisMinorTickMark) {
    }

    /// <summary>
    /// Creates chart layout metadata with explicit legend overlay behavior.
    /// </summary>
    /// <param name="overlayLegend">Whether a legend should be drawn over the plot instead of reserving layout space.</param>
    /// <param name="seriesLegendWidthRatio">Maximum chart-width ratio reserved for series legends.</param>
    /// <param name="categoryLegendWidthRatio">Maximum chart-width ratio reserved for category legends such as pie slices.</param>
    /// <param name="legendRowHeight">Legend row height.</param>
    /// <param name="legendSwatchSize">Legend color swatch size.</param>
    /// <param name="legendTextGap">Gap between a legend swatch and its label.</param>
    /// <param name="legendFontSize">Legend label font size.</param>
    /// <param name="legendFontFamily">Optional legend label font family.</param>
    /// <param name="axisLabelFontSize">Axis label font size.</param>
    /// <param name="axisTextFontFamily">Optional axis label font family and axis title fallback family.</param>
    /// <param name="categoryAxisLabelWidth">Maximum category-axis label width.</param>
    /// <param name="radarCategoryLabelWidth">Maximum radar category label width.</param>
    /// <param name="maximumCategoryAxisLabels">Maximum number of category-axis labels to render on cartesian charts.</param>
    /// <param name="maximumHorizontalCategoryAxisLabels">Maximum number of category-axis labels to render on horizontal bar charts.</param>
    /// <param name="maximumRadarCategoryLabels">Maximum number of category labels to render on radar charts.</param>
    /// <param name="preventLabelOverlap">Whether axis/category label stride should increase automatically to avoid obvious label overlap.</param>
    /// <param name="showLegend">Whether series or category legends should be rendered.</param>
    /// <param name="legendPosition">Preferred legend placement when legends are rendered.</param>
    /// <param name="showDataLabels">Whether point data labels should be rendered when supported by a chart family.</param>
    /// <param name="showDataLabelValues">Whether point data labels should include values.</param>
    /// <param name="showDataLabelPercentages">Whether point data labels should include percentages.</param>
    /// <param name="showDataLabelCategoryNames">Whether point data labels should include category names.</param>
    /// <param name="showDataLabelSeriesNames">Whether point data labels should include series names.</param>
    /// <param name="dataLabelSeparator">Separator used between enabled data label parts.</param>
    /// <param name="dataLabelFontSize">Data label font size.</param>
    /// <param name="dataLabelFontFamily">Optional data label font family.</param>
    /// <param name="dataLabelPosition">Preferred data label position when labels are rendered.</param>
    /// <param name="dataLabelNumberFormat">Optional numeric format for data-label values.</param>
    /// <param name="showMarkers">Whether point markers should be rendered for marker-capable chart families.</param>
    /// <param name="axisNumberFormat">Optional numeric format for value-axis labels.</param>
    /// <param name="categoryAxisTitle">Optional category or horizontal axis title.</param>
    /// <param name="valueAxisTitle">Optional value or vertical axis title.</param>
    /// <param name="horizontalAxisNumberFormat">Optional numeric format for horizontal value-axis labels.</param>
    /// <param name="verticalAxisNumberFormat">Optional numeric format for vertical value-axis labels.</param>
    /// <param name="connectScatterPoints">Whether scatter points should be connected by series lines.</param>
    /// <param name="fillRadarSeries">Whether radar series polygons should be filled.</param>
    /// <param name="showCategoryAxis">Whether the category or horizontal axis should be rendered.</param>
    /// <param name="showValueAxis">Whether the value or vertical axis should be rendered.</param>
    /// <param name="showCategoryAxisLine">Whether the category or horizontal axis line should be rendered.</param>
    /// <param name="showValueAxisLine">Whether the value or vertical axis line should be rendered.</param>
    /// <param name="showCategoryAxisLabels">Whether category or horizontal tick labels should be rendered.</param>
    /// <param name="showValueAxisLabels">Whether value or vertical tick labels should be rendered.</param>
    /// <param name="horizontalAxisTickLabelPosition">Preferred tick-label side for the physical horizontal axis.</param>
    /// <param name="verticalAxisTickLabelPosition">Preferred tick-label side for the physical vertical axis.</param>
    /// <param name="horizontalAxisCrossingPosition">Physical crossing side for the horizontal axis.</param>
    /// <param name="verticalAxisCrossingPosition">Physical crossing side for the vertical axis.</param>
    /// <param name="reverseCategoryAxis">Whether category-axis slots should be rendered in reverse order.</param>
    /// <param name="overlayTitle">Whether the title should overlay the plot instead of reserving layout space.</param>
    /// <param name="titleTopPadding">Top padding before the chart title inside the chart canvas.</param>
    /// <param name="legendFontStyle">Optional legend label font style.</param>
    /// <param name="axisTextFontStyle">Optional axis label font style and axis title fallback style.</param>
    /// <param name="dataLabelFontStyle">Optional data label font style.</param>
    /// <param name="axisTitleFontSize">Optional axis title font size.</param>
    /// <param name="axisTitleFontFamily">Optional axis title font family.</param>
    /// <param name="axisTitleFontStyle">Optional axis title font style.</param>
    /// <param name="horizontalAxisDisplayUnitDivisor">Optional divisor applied to horizontal value-axis labels.</param>
    /// <param name="horizontalAxisDisplayUnitLabel">Optional display-unit label shown for horizontal value axes.</param>
    /// <param name="verticalAxisDisplayUnitDivisor">Optional divisor applied to vertical value-axis labels.</param>
    /// <param name="verticalAxisDisplayUnitLabel">Optional display-unit label shown for vertical value axes.</param>
    /// <param name="horizontalAxisMinimum">Optional minimum for horizontal value axes.</param>
    /// <param name="horizontalAxisMaximum">Optional maximum for horizontal value axes.</param>
    /// <param name="horizontalAxisMajorUnit">Optional major tick/grid unit for horizontal value axes.</param>
    /// <param name="horizontalAxisMinorUnit">Optional minor tick/grid unit for horizontal value axes.</param>
    /// <param name="verticalAxisMinimum">Optional minimum for vertical value axes.</param>
    /// <param name="verticalAxisMaximum">Optional maximum for vertical value axes.</param>
    /// <param name="verticalAxisMajorUnit">Optional major tick/grid unit for vertical value axes.</param>
    /// <param name="verticalAxisMinorUnit">Optional minor tick/grid unit for vertical value axes.</param>
    /// <param name="horizontalAxisMajorTickMark">Optional major tick mark placement for the horizontal axis.</param>
    /// <param name="verticalAxisMajorTickMark">Optional major tick mark placement for the vertical axis.</param>
    /// <param name="horizontalAxisMinorTickMark">Optional minor tick mark placement for the horizontal axis.</param>
    /// <param name="verticalAxisMinorTickMark">Optional minor tick mark placement for the vertical axis.</param>
    public OfficeChartLayout(
        bool overlayLegend,
        double? seriesLegendWidthRatio = null,
        double? categoryLegendWidthRatio = null,
        double? legendRowHeight = null,
        double? legendSwatchSize = null,
        double? legendTextGap = null,
        double? legendFontSize = null,
        string? legendFontFamily = null,
        double? axisLabelFontSize = null,
        string? axisTextFontFamily = null,
        double? categoryAxisLabelWidth = null,
        double? radarCategoryLabelWidth = null,
        int? maximumCategoryAxisLabels = null,
        int? maximumHorizontalCategoryAxisLabels = null,
        int? maximumRadarCategoryLabels = null,
        bool preventLabelOverlap = true,
        bool showLegend = true,
        OfficeChartLegendPosition legendPosition = OfficeChartLegendPosition.Right,
        bool showDataLabels = false,
        bool showDataLabelValues = false,
        bool showDataLabelPercentages = false,
        bool showDataLabelCategoryNames = false,
        bool showDataLabelSeriesNames = false,
        string? dataLabelSeparator = null,
        double? dataLabelFontSize = null,
        string? dataLabelFontFamily = null,
        OfficeChartDataLabelPosition dataLabelPosition = OfficeChartDataLabelPosition.Auto,
        string? dataLabelNumberFormat = null,
        bool showMarkers = true,
        string? axisNumberFormat = null,
        string? categoryAxisTitle = null,
        string? valueAxisTitle = null,
        string? horizontalAxisNumberFormat = null,
        string? verticalAxisNumberFormat = null,
        bool connectScatterPoints = true,
        bool fillRadarSeries = true,
        bool showCategoryAxis = true,
        bool showValueAxis = true,
        bool showCategoryAxisLine = true,
        bool showValueAxisLine = true,
        bool showCategoryAxisLabels = true,
        bool showValueAxisLabels = true,
        OfficeChartAxisTickLabelPosition horizontalAxisTickLabelPosition = OfficeChartAxisTickLabelPosition.NextTo,
        OfficeChartAxisTickLabelPosition verticalAxisTickLabelPosition = OfficeChartAxisTickLabelPosition.NextTo,
        OfficeChartAxisCrossingPosition horizontalAxisCrossingPosition = OfficeChartAxisCrossingPosition.AutoZero,
        OfficeChartAxisCrossingPosition verticalAxisCrossingPosition = OfficeChartAxisCrossingPosition.AutoZero,
        bool reverseCategoryAxis = false,
        bool overlayTitle = false,
        double? titleTopPadding = null,
        OfficeFontStyle? legendFontStyle = null,
        OfficeFontStyle? axisTextFontStyle = null,
        OfficeFontStyle? dataLabelFontStyle = null,
        double? axisTitleFontSize = null,
        string? axisTitleFontFamily = null,
        OfficeFontStyle? axisTitleFontStyle = null,
        double? horizontalAxisDisplayUnitDivisor = null,
        string? horizontalAxisDisplayUnitLabel = null,
        double? verticalAxisDisplayUnitDivisor = null,
        string? verticalAxisDisplayUnitLabel = null,
        double? horizontalAxisMinimum = null,
        double? horizontalAxisMaximum = null,
        double? horizontalAxisMajorUnit = null,
        double? horizontalAxisMinorUnit = null,
        double? verticalAxisMinimum = null,
        double? verticalAxisMaximum = null,
        double? verticalAxisMajorUnit = null,
        double? verticalAxisMinorUnit = null,
        OfficeChartAxisTickMark horizontalAxisMajorTickMark = OfficeChartAxisTickMark.None,
        OfficeChartAxisTickMark verticalAxisMajorTickMark = OfficeChartAxisTickMark.None,
        OfficeChartAxisTickMark horizontalAxisMinorTickMark = OfficeChartAxisTickMark.None,
        OfficeChartAxisTickMark verticalAxisMinorTickMark = OfficeChartAxisTickMark.None) {
        SeriesLegendWidthRatio = ValidateRatio(seriesLegendWidthRatio ?? 0.34D, nameof(seriesLegendWidthRatio));
        CategoryLegendWidthRatio = ValidateRatio(categoryLegendWidthRatio ?? 0.38D, nameof(categoryLegendWidthRatio));
        LegendRowHeight = ValidatePositiveFinite(legendRowHeight ?? 12D, nameof(legendRowHeight));
        LegendSwatchSize = ValidatePositiveFinite(legendSwatchSize ?? 6D, nameof(legendSwatchSize));
        LegendTextGap = ValidatePositiveFinite(legendTextGap ?? 4D, nameof(legendTextGap));
        LegendFontSize = ValidatePositiveFinite(legendFontSize ?? 7.2D, nameof(legendFontSize));
        LegendFontFamily = string.IsNullOrWhiteSpace(legendFontFamily) ? null : legendFontFamily;
        LegendFontStyle = legendFontStyle;
        AxisLabelFontSize = ValidatePositiveFinite(axisLabelFontSize ?? 6.8D, nameof(axisLabelFontSize));
        AxisTextFontFamily = string.IsNullOrWhiteSpace(axisTextFontFamily) ? null : axisTextFontFamily;
        AxisTextFontStyle = axisTextFontStyle;
        AxisTitleFontSize = axisTitleFontSize.HasValue ? ValidatePositiveFinite(axisTitleFontSize.Value, nameof(axisTitleFontSize)) : null;
        AxisTitleFontFamily = string.IsNullOrWhiteSpace(axisTitleFontFamily) ? null : axisTitleFontFamily;
        AxisTitleFontStyle = axisTitleFontStyle;
        CategoryAxisLabelWidth = ValidatePositiveFinite(categoryAxisLabelWidth ?? 54D, nameof(categoryAxisLabelWidth));
        RadarCategoryLabelWidth = ValidatePositiveFinite(radarCategoryLabelWidth ?? 42D, nameof(radarCategoryLabelWidth));
        MaximumCategoryAxisLabels = ValidatePositive(maximumCategoryAxisLabels ?? 6, nameof(maximumCategoryAxisLabels));
        MaximumHorizontalCategoryAxisLabels = ValidatePositive(maximumHorizontalCategoryAxisLabels ?? 7, nameof(maximumHorizontalCategoryAxisLabels));
        MaximumRadarCategoryLabels = ValidatePositive(maximumRadarCategoryLabels ?? 8, nameof(maximumRadarCategoryLabels));
        PreventLabelOverlap = preventLabelOverlap;
        ShowLegend = showLegend;
        LegendPosition = legendPosition;
        OverlayLegend = overlayLegend;
        ShowDataLabels = showDataLabels;
        ShowDataLabelValues = showDataLabelValues;
        ShowDataLabelPercentages = showDataLabelPercentages;
        ShowDataLabelCategoryNames = showDataLabelCategoryNames;
        ShowDataLabelSeriesNames = showDataLabelSeriesNames;
        DataLabelSeparator = string.IsNullOrEmpty(dataLabelSeparator) ? "; " : dataLabelSeparator!;
        DataLabelFontSize = ValidatePositiveFinite(dataLabelFontSize ?? 7D, nameof(dataLabelFontSize));
        DataLabelFontFamily = string.IsNullOrWhiteSpace(dataLabelFontFamily) ? null : dataLabelFontFamily;
        DataLabelFontStyle = dataLabelFontStyle;
        DataLabelPosition = dataLabelPosition;
        DataLabelNumberFormat = NormalizeNumberFormat(dataLabelNumberFormat);
        ShowMarkers = showMarkers;
        AxisNumberFormat = NormalizeNumberFormat(axisNumberFormat);
        HorizontalAxisNumberFormat = string.IsNullOrWhiteSpace(horizontalAxisNumberFormat) ? AxisNumberFormat : NormalizeNumberFormat(horizontalAxisNumberFormat);
        VerticalAxisNumberFormat = string.IsNullOrWhiteSpace(verticalAxisNumberFormat) ? AxisNumberFormat : NormalizeNumberFormat(verticalAxisNumberFormat);
        HorizontalAxisDisplayUnitDivisor = horizontalAxisDisplayUnitDivisor.HasValue ? ValidatePositiveFinite(horizontalAxisDisplayUnitDivisor.Value, nameof(horizontalAxisDisplayUnitDivisor)) : null;
        HorizontalAxisDisplayUnitLabel = string.IsNullOrWhiteSpace(horizontalAxisDisplayUnitLabel) ? null : horizontalAxisDisplayUnitLabel;
        VerticalAxisDisplayUnitDivisor = verticalAxisDisplayUnitDivisor.HasValue ? ValidatePositiveFinite(verticalAxisDisplayUnitDivisor.Value, nameof(verticalAxisDisplayUnitDivisor)) : null;
        VerticalAxisDisplayUnitLabel = string.IsNullOrWhiteSpace(verticalAxisDisplayUnitLabel) ? null : verticalAxisDisplayUnitLabel;
        HorizontalAxisMinimum = horizontalAxisMinimum.HasValue ? ValidateFinite(horizontalAxisMinimum.Value, nameof(horizontalAxisMinimum)) : null;
        HorizontalAxisMaximum = horizontalAxisMaximum.HasValue ? ValidateFinite(horizontalAxisMaximum.Value, nameof(horizontalAxisMaximum)) : null;
        HorizontalAxisMajorUnit = horizontalAxisMajorUnit.HasValue ? ValidatePositiveFinite(horizontalAxisMajorUnit.Value, nameof(horizontalAxisMajorUnit)) : null;
        HorizontalAxisMinorUnit = horizontalAxisMinorUnit.HasValue ? ValidatePositiveFinite(horizontalAxisMinorUnit.Value, nameof(horizontalAxisMinorUnit)) : null;
        VerticalAxisMinimum = verticalAxisMinimum.HasValue ? ValidateFinite(verticalAxisMinimum.Value, nameof(verticalAxisMinimum)) : null;
        VerticalAxisMaximum = verticalAxisMaximum.HasValue ? ValidateFinite(verticalAxisMaximum.Value, nameof(verticalAxisMaximum)) : null;
        VerticalAxisMajorUnit = verticalAxisMajorUnit.HasValue ? ValidatePositiveFinite(verticalAxisMajorUnit.Value, nameof(verticalAxisMajorUnit)) : null;
        VerticalAxisMinorUnit = verticalAxisMinorUnit.HasValue ? ValidatePositiveFinite(verticalAxisMinorUnit.Value, nameof(verticalAxisMinorUnit)) : null;
        HorizontalAxisMajorTickMark = horizontalAxisMajorTickMark;
        VerticalAxisMajorTickMark = verticalAxisMajorTickMark;
        HorizontalAxisMinorTickMark = horizontalAxisMinorTickMark;
        VerticalAxisMinorTickMark = verticalAxisMinorTickMark;
        CategoryAxisTitle = string.IsNullOrWhiteSpace(categoryAxisTitle) ? null : categoryAxisTitle;
        ValueAxisTitle = string.IsNullOrWhiteSpace(valueAxisTitle) ? null : valueAxisTitle;
        ConnectScatterPoints = connectScatterPoints;
        FillRadarSeries = fillRadarSeries;
        ShowCategoryAxis = showCategoryAxis;
        ShowValueAxis = showValueAxis;
        ShowCategoryAxisLine = showCategoryAxis && showCategoryAxisLine;
        ShowValueAxisLine = showValueAxis && showValueAxisLine;
        HorizontalAxisTickLabelPosition = horizontalAxisTickLabelPosition;
        VerticalAxisTickLabelPosition = verticalAxisTickLabelPosition;
        HorizontalAxisCrossingPosition = horizontalAxisCrossingPosition;
        VerticalAxisCrossingPosition = verticalAxisCrossingPosition;
        ReverseCategoryAxis = reverseCategoryAxis;
        ShowCategoryAxisLabels = showCategoryAxis && showCategoryAxisLabels;
        ShowValueAxisLabels = showValueAxis && showValueAxisLabels;
        OverlayTitle = overlayTitle;
        TitleTopPadding = ValidateNonNegativeFinite(titleTopPadding ?? 5D, nameof(titleTopPadding));
    }

    /// <summary>Default premium OfficeIMO chart layout.</summary>
    public static OfficeChartLayout Default => DefaultLayout;

    /// <summary>Maximum chart-width ratio reserved for series legends.</summary>
    public double SeriesLegendWidthRatio { get; }

    /// <summary>Maximum chart-width ratio reserved for category legends such as pie slices.</summary>
    public double CategoryLegendWidthRatio { get; }

    /// <summary>Legend row height.</summary>
    public double LegendRowHeight { get; }

    /// <summary>Legend color swatch size.</summary>
    public double LegendSwatchSize { get; }

    /// <summary>Gap between a legend swatch and its label.</summary>
    public double LegendTextGap { get; }

    /// <summary>Legend label font size.</summary>
    public double LegendFontSize { get; }

    /// <summary>Optional legend label font family.</summary>
    public string? LegendFontFamily { get; }

    /// <summary>Optional legend label font style.</summary>
    public OfficeFontStyle? LegendFontStyle { get; }

    /// <summary>Axis label font size.</summary>
    public double AxisLabelFontSize { get; }

    /// <summary>Optional axis label font family and axis title fallback family.</summary>
    public string? AxisTextFontFamily { get; }

    /// <summary>Optional axis label font style and axis title fallback style.</summary>
    public OfficeFontStyle? AxisTextFontStyle { get; }

    /// <summary>Optional axis title font size.</summary>
    public double? AxisTitleFontSize { get; }

    /// <summary>Optional axis title font family.</summary>
    public string? AxisTitleFontFamily { get; }

    /// <summary>Optional axis title font style.</summary>
    public OfficeFontStyle? AxisTitleFontStyle { get; }

    /// <summary>Maximum category-axis label width.</summary>
    public double CategoryAxisLabelWidth { get; }

    /// <summary>Maximum radar category label width.</summary>
    public double RadarCategoryLabelWidth { get; }

    /// <summary>Maximum number of category-axis labels to render on cartesian charts.</summary>
    public int MaximumCategoryAxisLabels { get; }

    /// <summary>Maximum number of category-axis labels to render on horizontal bar charts.</summary>
    public int MaximumHorizontalCategoryAxisLabels { get; }

    /// <summary>Maximum number of category labels to render on radar charts.</summary>
    public int MaximumRadarCategoryLabels { get; }

    /// <summary>Whether axis/category label stride should increase automatically to avoid obvious label overlap.</summary>
    public bool PreventLabelOverlap { get; }

    /// <summary>Whether series or category legends should be rendered.</summary>
    public bool ShowLegend { get; }

    /// <summary>Preferred legend placement when legends are rendered.</summary>
    public OfficeChartLegendPosition LegendPosition { get; }

    /// <summary>Whether a legend should be drawn over the plot instead of reserving layout space.</summary>
    public bool OverlayLegend { get; }

    /// <summary>Whether point data labels should be rendered when supported by a chart family.</summary>
    public bool ShowDataLabels { get; }

    /// <summary>Whether point data labels should include values.</summary>
    public bool ShowDataLabelValues { get; }

    /// <summary>Whether point data labels should include percentages.</summary>
    public bool ShowDataLabelPercentages { get; }

    /// <summary>Whether point data labels should include category names.</summary>
    public bool ShowDataLabelCategoryNames { get; }

    /// <summary>Whether point data labels should include series names.</summary>
    public bool ShowDataLabelSeriesNames { get; }

    /// <summary>Separator used between enabled data label parts.</summary>
    public string DataLabelSeparator { get; }

    /// <summary>Data label font size.</summary>
    public double DataLabelFontSize { get; }

    /// <summary>Optional data label font family.</summary>
    public string? DataLabelFontFamily { get; }

    /// <summary>Optional data label font style.</summary>
    public OfficeFontStyle? DataLabelFontStyle { get; }

    /// <summary>Preferred data label position when labels are rendered.</summary>
    public OfficeChartDataLabelPosition DataLabelPosition { get; }

    /// <summary>Optional numeric format for data-label values.</summary>
    public string? DataLabelNumberFormat { get; }

    /// <summary>Optional zero-based series indexes that should render data labels; null means every series may render labels.</summary>
    public IReadOnlyCollection<int>? DataLabelSeriesIndexes { get; set; }

    /// <summary>Optional zero-based point indexes that should render data labels per series; null means every point may render labels.</summary>
    public IReadOnlyDictionary<int, IReadOnlyCollection<int>>? DataLabelPointIndexes { get; set; }

    /// <summary>Optional zero-based point indexes that should suppress data labels per series.</summary>
    public IReadOnlyDictionary<int, IReadOnlyCollection<int>>? HiddenDataLabelPointIndexes { get; set; }

    /// <summary>Optional zero-based category or slice legend indexes that should be suppressed for category legends.</summary>
    public IReadOnlyCollection<int>? HiddenCategoryLegendIndexes { get; set; }

    /// <summary>Whether point markers should be rendered for marker-capable chart families.</summary>
    public bool ShowMarkers { get; }

    /// <summary>Optional numeric format for value-axis labels.</summary>
    public string? AxisNumberFormat { get; }

    /// <summary>Optional numeric format for horizontal value-axis labels.</summary>
    public string? HorizontalAxisNumberFormat { get; }

    /// <summary>Optional numeric format for vertical value-axis labels.</summary>
    public string? VerticalAxisNumberFormat { get; }

    /// <summary>Optional divisor applied to horizontal value-axis labels.</summary>
    public double? HorizontalAxisDisplayUnitDivisor { get; }

    /// <summary>Optional display-unit label shown for horizontal value axes.</summary>
    public string? HorizontalAxisDisplayUnitLabel { get; }

    /// <summary>Optional divisor applied to vertical value-axis labels.</summary>
    public double? VerticalAxisDisplayUnitDivisor { get; }

    /// <summary>Optional display-unit label shown for vertical value axes.</summary>
    public string? VerticalAxisDisplayUnitLabel { get; }

    /// <summary>Optional minimum for horizontal value axes.</summary>
    public double? HorizontalAxisMinimum { get; }

    /// <summary>Optional maximum for horizontal value axes.</summary>
    public double? HorizontalAxisMaximum { get; }

    /// <summary>Optional major tick/grid unit for horizontal value axes.</summary>
    public double? HorizontalAxisMajorUnit { get; }

    /// <summary>Optional minor tick/grid unit for horizontal value axes.</summary>
    public double? HorizontalAxisMinorUnit { get; }

    /// <summary>Optional minimum for vertical value axes.</summary>
    public double? VerticalAxisMinimum { get; }

    /// <summary>Optional maximum for vertical value axes.</summary>
    public double? VerticalAxisMaximum { get; }

    /// <summary>Optional major tick/grid unit for vertical value axes.</summary>
    public double? VerticalAxisMajorUnit { get; }

    /// <summary>Optional minor tick/grid unit for vertical value axes.</summary>
    public double? VerticalAxisMinorUnit { get; }

    /// <summary>Major tick mark placement for the horizontal axis.</summary>
    public OfficeChartAxisTickMark HorizontalAxisMajorTickMark { get; }

    /// <summary>Major tick mark placement for the vertical axis.</summary>
    public OfficeChartAxisTickMark VerticalAxisMajorTickMark { get; }

    /// <summary>Minor tick mark placement for the horizontal axis.</summary>
    public OfficeChartAxisTickMark HorizontalAxisMinorTickMark { get; }

    /// <summary>Minor tick mark placement for the vertical axis.</summary>
    public OfficeChartAxisTickMark VerticalAxisMinorTickMark { get; }

    /// <summary>Optional category or horizontal axis title.</summary>
    public string? CategoryAxisTitle { get; }

    /// <summary>Optional value or vertical axis title.</summary>
    public string? ValueAxisTitle { get; }

    /// <summary>Whether scatter points should be connected by series lines.</summary>
    public bool ConnectScatterPoints { get; }

    /// <summary>Whether radar series polygons should be filled.</summary>
    public bool FillRadarSeries { get; }

    /// <summary>Whether the category or horizontal axis should be rendered.</summary>
    public bool ShowCategoryAxis { get; }

    /// <summary>Whether the value or vertical axis should be rendered.</summary>
    public bool ShowValueAxis { get; }

    /// <summary>Whether the category or horizontal axis line should be rendered.</summary>
    public bool ShowCategoryAxisLine { get; }

    /// <summary>Whether the value or vertical axis line should be rendered.</summary>
    public bool ShowValueAxisLine { get; }

    /// <summary>Whether category or horizontal tick labels should be rendered.</summary>
    public bool ShowCategoryAxisLabels { get; }

    /// <summary>Whether value or vertical tick labels should be rendered.</summary>
    public bool ShowValueAxisLabels { get; }

    /// <summary>Preferred tick-label side for the physical horizontal axis.</summary>
    public OfficeChartAxisTickLabelPosition HorizontalAxisTickLabelPosition { get; }

    /// <summary>Preferred tick-label side for the physical vertical axis.</summary>
    public OfficeChartAxisTickLabelPosition VerticalAxisTickLabelPosition { get; }

    /// <summary>Physical crossing side for the horizontal axis.</summary>
    public OfficeChartAxisCrossingPosition HorizontalAxisCrossingPosition { get; }

    /// <summary>Physical crossing side for the vertical axis.</summary>
    public OfficeChartAxisCrossingPosition VerticalAxisCrossingPosition { get; }

    /// <summary>Whether category-axis slots should be rendered in reverse order.</summary>
    public bool ReverseCategoryAxis { get; }

    /// <summary>Whether the chart title should overlay the plot area instead of reserving a title band.</summary>
    public bool OverlayTitle { get; }

    /// <summary>Top padding before the chart title inside the chart canvas.</summary>
    public double TitleTopPadding { get; }

    private static double ValidateRatio(double value, string paramName) {
        ValidatePositiveFinite(value, paramName);
        if (value > 0.75D) {
            throw new ArgumentOutOfRangeException(paramName, "Chart legend width ratios must be less than or equal to 0.75.");
        }

        return value;
    }

    private static double ValidatePositiveFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) {
            throw new ArgumentOutOfRangeException(paramName, "Chart layout values must be finite positive numbers.");
        }

        return value;
    }

    private static double ValidateFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, "Chart layout values must be finite numbers.");
        }

        return value;
    }

    private static double ValidateNonNegativeFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0D) {
            throw new ArgumentOutOfRangeException(paramName, "Chart layout values must be finite non-negative numbers.");
        }

        return value;
    }

    private static int ValidatePositive(int value, string paramName) {
        if (value <= 0) {
            throw new ArgumentOutOfRangeException(paramName, "Chart layout counts must be positive.");
        }

        return value;
    }

    private static string? NormalizeNumberFormat(string? value) {
        if (string.IsNullOrWhiteSpace(value)) {
            return null;
        }

        string normalized = value!.Trim();
        return normalized.Length <= MaxNumberFormatLength ? normalized : null;
    }
}
