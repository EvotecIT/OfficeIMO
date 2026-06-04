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
    public OfficeChartSeries(string name, IEnumerable<double> values) {
        if (values == null) {
            throw new ArgumentNullException(nameof(values));
        }

        Name = name ?? string.Empty;
        Values = new ReadOnlyCollection<double>(new List<double>(values));
    }

    /// <summary>Series display name.</summary>
    public string Name { get; }

    /// <summary>Series values aligned with chart categories.</summary>
    public IReadOnlyList<double> Values { get; }
}
