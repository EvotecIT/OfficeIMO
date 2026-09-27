using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free chart snapshot that can be rendered by shared OfficeIMO visual engines.
/// </summary>
public sealed class OfficeChartSnapshot {
    /// <summary>
    /// Initializes a chart snapshot for rendering.
    /// </summary>
    /// <param name="name">Source shape or drawing name.</param>
    /// <param name="title">Optional display title.</param>
    /// <param name="chartKind">Supported chart family.</param>
    /// <param name="data">Chart category and series data.</param>
    /// <param name="widthPoints">Requested render width in points.</param>
    /// <param name="heightPoints">Requested render height in points.</param>
    /// <param name="style">Optional shared chart style metadata.</param>
    /// <param name="layout">Optional shared chart layout metadata.</param>
    public OfficeChartSnapshot(string name, string? title, OfficeChartKind chartKind,
        OfficeChartData data, double widthPoints, double heightPoints,
        OfficeChartStyle? style = null, OfficeChartLayout? layout = null)
        : this(name, title, chartKind, data, widthPoints, heightPoints,
            style, layout, 100D, OfficeChartBubbleSizeMode.Area) {
    }

    /// <summary>
    /// Initializes a bubble-aware chart snapshot for rendering with default style and layout.
    /// </summary>
    /// <param name="name">Source shape or drawing name.</param>
    /// <param name="title">Optional display title.</param>
    /// <param name="chartKind">Supported chart family.</param>
    /// <param name="data">Chart category and series data.</param>
    /// <param name="widthPoints">Requested render width in points.</param>
    /// <param name="heightPoints">Requested render height in points.</param>
    /// <param name="bubbleScalePercent">Bubble diameter scale as a percentage from zero through 300.</param>
    /// <param name="bubbleSizeMode">Whether bubble values represent area or width.</param>
    public OfficeChartSnapshot(string name, string? title, OfficeChartKind chartKind,
        OfficeChartData data, double widthPoints, double heightPoints,
        double bubbleScalePercent,
        OfficeChartBubbleSizeMode bubbleSizeMode = OfficeChartBubbleSizeMode.Area)
        : this(name, title, chartKind, data, widthPoints, heightPoints,
            null, null, bubbleScalePercent, bubbleSizeMode) {
    }

    /// <summary>
    /// Initializes a bubble-aware chart snapshot for rendering.
    /// </summary>
    /// <param name="name">Source shape or drawing name.</param>
    /// <param name="title">Optional display title.</param>
    /// <param name="chartKind">Supported chart family.</param>
    /// <param name="data">Chart category and series data.</param>
    /// <param name="widthPoints">Requested render width in points.</param>
    /// <param name="heightPoints">Requested render height in points.</param>
    /// <param name="style">Optional shared chart style metadata.</param>
    /// <param name="layout">Optional shared chart layout metadata.</param>
    /// <param name="bubbleScalePercent">Bubble diameter scale as a percentage from zero through 300.</param>
    /// <param name="bubbleSizeMode">Whether bubble values represent area or width.</param>
    public OfficeChartSnapshot(string name, string? title, OfficeChartKind chartKind,
        OfficeChartData data, double widthPoints, double heightPoints,
        OfficeChartStyle? style, OfficeChartLayout? layout,
        double bubbleScalePercent, OfficeChartBubbleSizeMode bubbleSizeMode)
        : this(name, title, chartKind, data, widthPoints, heightPoints, style, layout,
            bubbleScalePercent, bubbleSizeMode, OfficeChartRadialLayout.Default) {
    }

    /// <summary>Creates a chart snapshot with explicit radial geometry and default bubble sizing.</summary>
    public OfficeChartSnapshot(string name, string? title, OfficeChartKind chartKind,
        OfficeChartData data, double widthPoints, double heightPoints,
        OfficeChartStyle? style, OfficeChartLayout? layout, OfficeChartRadialLayout radialLayout)
        : this(name, title, chartKind, data, widthPoints, heightPoints, style, layout,
            100D, OfficeChartBubbleSizeMode.Area, radialLayout) {
    }

    /// <summary>Creates a chart snapshot with explicit bubble sizing and radial geometry.</summary>
    public OfficeChartSnapshot(string name, string? title, OfficeChartKind chartKind,
        OfficeChartData data, double widthPoints, double heightPoints,
        OfficeChartStyle? style, OfficeChartLayout? layout,
        double bubbleScalePercent, OfficeChartBubbleSizeMode bubbleSizeMode, OfficeChartRadialLayout radialLayout) {
        if (data == null) {
            throw new ArgumentNullException(nameof(data));
        }

        ValidatePositiveFinite(widthPoints, nameof(widthPoints));
        ValidatePositiveFinite(heightPoints, nameof(heightPoints));
        if (double.IsNaN(bubbleScalePercent) || double.IsInfinity(bubbleScalePercent) ||
            bubbleScalePercent < 0D || bubbleScalePercent > 300D) {
            throw new ArgumentOutOfRangeException(nameof(bubbleScalePercent),
                "Bubble scale must be a finite percentage from zero through 300.");
        }
        if (!Enum.IsDefined(typeof(OfficeChartBubbleSizeMode), bubbleSizeMode)) {
            throw new ArgumentOutOfRangeException(nameof(bubbleSizeMode));
        }
        for (int seriesIndex = 0; seriesIndex < data.Series.Count; seriesIndex++) {
            OfficeChartSeries series = data.Series[seriesIndex];
            OfficeChartKind effectiveKind = series.RenderKind ?? chartKind;
            if (effectiveKind == OfficeChartKind.Bubble &&
                chartKind != OfficeChartKind.Bubble &&
                chartKind != OfficeChartKind.Scatter) {
                throw new ArgumentException(
                    "Bubble series require a bubble or scatter snapshot with numeric axes.",
                    nameof(data));
            }
            if (effectiveKind == OfficeChartKind.Bubble &&
                (series.XValues == null || series.BubbleSizes == null ||
                 series.XValues.Count != series.Values.Count ||
                 series.BubbleSizes.Count != series.Values.Count)) {
                throw new ArgumentException(
                    "Bubble chart snapshots require matching X, Y, and size values for every bubble series.",
                    nameof(data));
            }
        }

        Name = name ?? string.Empty;
        Title = title;
        ChartKind = chartKind;
        Data = data;
        WidthPoints = widthPoints;
        HeightPoints = heightPoints;
        Style = style ?? OfficeChartStyle.Default;
        Layout = layout ?? OfficeChartLayout.Default;
        BubbleScalePercent = bubbleScalePercent;
        BubbleSizeMode = bubbleSizeMode;
        RadialLayout = radialLayout ?? throw new ArgumentNullException(nameof(radialLayout));
    }

    /// <summary>Source shape or drawing name.</summary>
    public string Name { get; }

    /// <summary>Optional display title.</summary>
    public string? Title { get; }

    /// <summary>Supported chart family.</summary>
    public OfficeChartKind ChartKind { get; }

    /// <summary>Chart category and series data.</summary>
    public OfficeChartData Data { get; }

    /// <summary>Requested render width in points.</summary>
    public double WidthPoints { get; }

    /// <summary>Requested render height in points.</summary>
    public double HeightPoints { get; }

    /// <summary>Shared chart style metadata.</summary>
    public OfficeChartStyle Style { get; }

    /// <summary>Shared chart layout metadata.</summary>
    public OfficeChartLayout Layout { get; }

    /// <summary>Bubble diameter scale as a percentage from zero through 300.</summary>
    public double BubbleScalePercent { get; }

    /// <summary>Whether bubble values represent area or width.</summary>
    public OfficeChartBubbleSizeMode BubbleSizeMode { get; }

    /// <summary>Pie rotation and doughnut hole geometry.</summary>
    public OfficeChartRadialLayout RadialLayout { get; }

    /// <summary>Creates a snapshot at a different render size while retaining all chart data and presentation metadata.</summary>
    /// <param name="widthPoints">Finite positive render width in points.</param>
    /// <param name="heightPoints">Finite positive render height in points.</param>
    public OfficeChartSnapshot WithSize(double widthPoints, double heightPoints) =>
        new OfficeChartSnapshot(Name, Title, ChartKind, Data, widthPoints, heightPoints,
            Style, Layout, BubbleScalePercent, BubbleSizeMode, RadialLayout);

    private static void ValidatePositiveFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) {
            throw new ArgumentOutOfRangeException(paramName, "Chart snapshot dimensions must be finite positive numbers.");
        }
    }
}
