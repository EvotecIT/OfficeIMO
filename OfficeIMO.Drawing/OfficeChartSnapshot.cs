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
    /// <param name="bubbleScalePercent">Bubble diameter scale as a percentage from zero through 300.</param>
    /// <param name="bubbleSizeMode">Whether bubble values represent area or width.</param>
    public OfficeChartSnapshot(string name, string? title, OfficeChartKind chartKind,
        OfficeChartData data, double widthPoints, double heightPoints,
        OfficeChartStyle? style = null, OfficeChartLayout? layout = null,
        double bubbleScalePercent = 100D,
        OfficeChartBubbleSizeMode bubbleSizeMode = OfficeChartBubbleSizeMode.Area) {
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

    private static void ValidatePositiveFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0D) {
            throw new ArgumentOutOfRangeException(paramName, "Chart snapshot dimensions must be finite positive numbers.");
        }
    }
}
