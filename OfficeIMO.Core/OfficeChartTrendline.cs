using System;

namespace OfficeIMO.Drawing;

/// <summary>
/// Describes a series trendline that chart renderers fit to the series values and draw over the plot.
/// </summary>
public sealed class OfficeChartTrendline {
    /// <summary>Creates a trendline description.</summary>
    /// <param name="type">Regression or averaging type.</param>
    /// <param name="order">Polynomial order from 2 to 6; required for <see cref="OfficeChartTrendlineType.Polynomial"/>.</param>
    /// <param name="period">Moving-average period of at least 2; required for <see cref="OfficeChartTrendlineType.MovingAverage"/>.</param>
    /// <param name="displayEquation">Whether the fitted equation is shown beside the line.</param>
    /// <param name="displayRSquared">Whether the coefficient of determination is shown beside the line.</param>
    /// <param name="color">Optional line colour; the series colour is used when omitted.</param>
    /// <param name="width">Optional line width in points.</param>
    /// <param name="intercept">Optional fixed y-intercept for linear, polynomial and exponential fits.</param>
    public OfficeChartTrendline(OfficeChartTrendlineType type, int? order = null, int? period = null,
        bool displayEquation = false, bool displayRSquared = false, OfficeColor? color = null, double? width = null,
        double? intercept = null) {
        if (!Enum.IsDefined(typeof(OfficeChartTrendlineType), type)) throw new ArgumentOutOfRangeException(nameof(type));
        if (type == OfficeChartTrendlineType.Polynomial && (order is null or < 2 or > 6))
            throw new ArgumentOutOfRangeException(nameof(order), "Polynomial trendlines need an order from 2 to 6.");
        if (type == OfficeChartTrendlineType.MovingAverage && (period is null or < 2))
            throw new ArgumentOutOfRangeException(nameof(period), "Moving-average trendlines need a period of at least 2.");
        if (width is double w && (double.IsNaN(w) || double.IsInfinity(w) || w <= 0D))
            throw new ArgumentOutOfRangeException(nameof(width), "Trendline width must be finite and greater than zero.");
        if (intercept is double i && (double.IsNaN(i) || double.IsInfinity(i)))
            throw new ArgumentOutOfRangeException(nameof(intercept), "Trendline intercept must be finite.");
        Type = type;
        Order = type == OfficeChartTrendlineType.Polynomial ? order : null;
        Period = type == OfficeChartTrendlineType.MovingAverage ? period : null;
        DisplayEquation = displayEquation && type != OfficeChartTrendlineType.MovingAverage;
        DisplayRSquared = displayRSquared && type != OfficeChartTrendlineType.MovingAverage;
        Color = color;
        Width = width;
        Intercept = intercept;
    }

    /// <summary>Regression or averaging type.</summary>
    public OfficeChartTrendlineType Type { get; }

    /// <summary>Polynomial order, or null for other types.</summary>
    public int? Order { get; }

    /// <summary>Moving-average period, or null for other types.</summary>
    public int? Period { get; }

    /// <summary>Whether the fitted equation is shown beside the line.</summary>
    public bool DisplayEquation { get; }

    /// <summary>Whether the coefficient of determination is shown beside the line.</summary>
    public bool DisplayRSquared { get; }

    /// <summary>Optional line colour.</summary>
    public OfficeColor? Color { get; }

    /// <summary>Optional line width in points.</summary>
    public double? Width { get; }

    /// <summary>Optional fixed y-intercept.</summary>
    public double? Intercept { get; }
}
