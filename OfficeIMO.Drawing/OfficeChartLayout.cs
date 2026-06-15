using System;

using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Reusable chart layout metadata shared by OfficeIMO chart renderers and format exporters.
/// </summary>
public sealed class OfficeChartLayout {
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
    /// <param name="axisLabelFontSize">Axis label font size.</param>
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
    /// <param name="overlayTitle">Whether the title should overlay the plot instead of reserving layout space.</param>
    public OfficeChartLayout(
        double? seriesLegendWidthRatio = null,
        double? categoryLegendWidthRatio = null,
        double? legendRowHeight = null,
        double? legendSwatchSize = null,
        double? legendTextGap = null,
        double? legendFontSize = null,
        double? axisLabelFontSize = null,
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
        bool overlayTitle = false)
        : this(
            overlayLegend: false,
            seriesLegendWidthRatio: seriesLegendWidthRatio,
            categoryLegendWidthRatio: categoryLegendWidthRatio,
            legendRowHeight: legendRowHeight,
            legendSwatchSize: legendSwatchSize,
            legendTextGap: legendTextGap,
            legendFontSize: legendFontSize,
            axisLabelFontSize: axisLabelFontSize,
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
            overlayTitle: overlayTitle) {
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
    /// <param name="axisLabelFontSize">Axis label font size.</param>
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
    /// <param name="overlayTitle">Whether the title should overlay the plot instead of reserving layout space.</param>
    public OfficeChartLayout(
        bool overlayLegend,
        double? seriesLegendWidthRatio = null,
        double? categoryLegendWidthRatio = null,
        double? legendRowHeight = null,
        double? legendSwatchSize = null,
        double? legendTextGap = null,
        double? legendFontSize = null,
        double? axisLabelFontSize = null,
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
        bool overlayTitle = false) {
        SeriesLegendWidthRatio = ValidateRatio(seriesLegendWidthRatio ?? 0.34D, nameof(seriesLegendWidthRatio));
        CategoryLegendWidthRatio = ValidateRatio(categoryLegendWidthRatio ?? 0.38D, nameof(categoryLegendWidthRatio));
        LegendRowHeight = ValidatePositiveFinite(legendRowHeight ?? 12D, nameof(legendRowHeight));
        LegendSwatchSize = ValidatePositiveFinite(legendSwatchSize ?? 6D, nameof(legendSwatchSize));
        LegendTextGap = ValidatePositiveFinite(legendTextGap ?? 4D, nameof(legendTextGap));
        LegendFontSize = ValidatePositiveFinite(legendFontSize ?? 7.2D, nameof(legendFontSize));
        AxisLabelFontSize = ValidatePositiveFinite(axisLabelFontSize ?? 6.8D, nameof(axisLabelFontSize));
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
        DataLabelPosition = dataLabelPosition;
        DataLabelNumberFormat = string.IsNullOrWhiteSpace(dataLabelNumberFormat) ? null : dataLabelNumberFormat;
        ShowMarkers = showMarkers;
        AxisNumberFormat = string.IsNullOrWhiteSpace(axisNumberFormat) ? null : axisNumberFormat;
        HorizontalAxisNumberFormat = string.IsNullOrWhiteSpace(horizontalAxisNumberFormat) ? AxisNumberFormat : horizontalAxisNumberFormat;
        VerticalAxisNumberFormat = string.IsNullOrWhiteSpace(verticalAxisNumberFormat) ? AxisNumberFormat : verticalAxisNumberFormat;
        CategoryAxisTitle = string.IsNullOrWhiteSpace(categoryAxisTitle) ? null : categoryAxisTitle;
        ValueAxisTitle = string.IsNullOrWhiteSpace(valueAxisTitle) ? null : valueAxisTitle;
        ConnectScatterPoints = connectScatterPoints;
        FillRadarSeries = fillRadarSeries;
        ShowCategoryAxis = showCategoryAxis;
        ShowValueAxis = showValueAxis;
        ShowCategoryAxisLine = showCategoryAxis && showCategoryAxisLine;
        ShowValueAxisLine = showValueAxis && showValueAxisLine;
        ShowCategoryAxisLabels = showCategoryAxis && showCategoryAxisLabels;
        ShowValueAxisLabels = showValueAxis && showValueAxisLabels;
        OverlayTitle = overlayTitle;
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

    /// <summary>Axis label font size.</summary>
    public double AxisLabelFontSize { get; }

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

    /// <summary>Whether point markers should be rendered for marker-capable chart families.</summary>
    public bool ShowMarkers { get; }

    /// <summary>Optional numeric format for value-axis labels.</summary>
    public string? AxisNumberFormat { get; }

    /// <summary>Optional numeric format for horizontal value-axis labels.</summary>
    public string? HorizontalAxisNumberFormat { get; }

    /// <summary>Optional numeric format for vertical value-axis labels.</summary>
    public string? VerticalAxisNumberFormat { get; }

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

    /// <summary>Whether the chart title should overlay the plot area instead of reserving a title band.</summary>
    public bool OverlayTitle { get; }

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

    private static int ValidatePositive(int value, string paramName) {
        if (value <= 0) {
            throw new ArgumentOutOfRangeException(paramName, "Chart layout counts must be positive.");
        }

        return value;
    }
}
