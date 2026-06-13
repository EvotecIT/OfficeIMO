using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free chart series values used by shared OfficeIMO visual renderers.
/// </summary>
public sealed class OfficeChartSeries {
    /// <summary>
    /// Initializes a chart series snapshot.
    /// </summary>
    /// <param name="name">Display name for the series.</param>
    /// <param name="values">Values aligned with the chart categories.</param>
    public OfficeChartSeries(string name, IEnumerable<double> values) : this(name, values, null, null, null) {
    }

    /// <summary>
    /// Initializes a chart series snapshot with optional numeric X-axis values for scatter charts.
    /// </summary>
    /// <param name="name">Display name for the series.</param>
    /// <param name="values">Values aligned with the chart categories or X-axis values.</param>
    /// <param name="xValues">Optional numeric X-axis values for this series.</param>
    public OfficeChartSeries(string name, IEnumerable<double> values, IEnumerable<double>? xValues) : this(name, values, xValues, null, null) {
    }

    /// <summary>
    /// Initializes a chart series snapshot with optional numeric X-axis values and source style metadata.
    /// </summary>
    /// <param name="name">Display name for the series.</param>
    /// <param name="values">Values aligned with the chart categories or X-axis values.</param>
    /// <param name="xValues">Optional numeric X-axis values for this series.</param>
    /// <param name="color">Optional source-defined series color.</param>
    public OfficeChartSeries(string name, IEnumerable<double> values, IEnumerable<double>? xValues, OfficeColor? color) : this(name, values, xValues, color, null) {
    }

    /// <summary>
    /// Initializes a chart series snapshot with optional source style metadata.
    /// </summary>
    /// <param name="name">Display name for the series.</param>
    /// <param name="values">Values aligned with the chart categories or X-axis values.</param>
    /// <param name="xValues">Optional numeric X-axis values for this series.</param>
    /// <param name="color">Optional source-defined series color.</param>
    /// <param name="pointColors">Optional source-defined colors aligned with individual values.</param>
    public OfficeChartSeries(string name, IEnumerable<double> values, IEnumerable<double>? xValues, OfficeColor? color, IEnumerable<OfficeColor?>? pointColors) {
        if (values == null) {
            throw new ArgumentNullException(nameof(values));
        }

        Name = name ?? string.Empty;
        Values = new ReadOnlyCollection<double>(new List<double>(values));
        if (xValues != null) {
            XValues = new ReadOnlyCollection<double>(new List<double>(xValues));
            if (XValues.Count != Values.Count) {
                throw new ArgumentException("Series X-axis values must match the number of series values.", nameof(xValues));
            }
        }

        Color = color;
        if (pointColors != null) {
            PointColors = new ReadOnlyCollection<OfficeColor?>(new List<OfficeColor?>(pointColors));
            if (PointColors.Count != Values.Count) {
                throw new ArgumentException("Series point colors must match the number of series values.", nameof(pointColors));
            }
        }
    }

    /// <summary>Series display name.</summary>
    public string Name { get; }

    /// <summary>Series values aligned with chart categories.</summary>
    public IReadOnlyList<double> Values { get; }

    /// <summary>Optional per-series numeric X-axis values for scatter charts.</summary>
    public IReadOnlyList<double>? XValues { get; }

    /// <summary>Optional source-defined series color.</summary>
    public OfficeColor? Color { get; }

    /// <summary>Optional source-defined colors aligned with individual series values.</summary>
    public IReadOnlyList<OfficeColor?>? PointColors { get; }
}
