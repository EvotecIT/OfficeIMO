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
    public OfficeChartSeries(string name, IEnumerable<double> values) : this(name, values, null) {
    }

    /// <summary>
    /// Initializes a chart series snapshot with optional numeric X-axis values for scatter charts.
    /// </summary>
    /// <param name="name">Display name for the series.</param>
    /// <param name="values">Values aligned with the chart categories or X-axis values.</param>
    /// <param name="xValues">Optional numeric X-axis values for this series.</param>
    public OfficeChartSeries(string name, IEnumerable<double> values, IEnumerable<double>? xValues) {
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
    }

    /// <summary>Series display name.</summary>
    public string Name { get; }

    /// <summary>Series values aligned with chart categories.</summary>
    public IReadOnlyList<double> Values { get; }

    /// <summary>Optional per-series numeric X-axis values for scatter charts.</summary>
    public IReadOnlyList<double>? XValues { get; }
}
