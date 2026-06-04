using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free chart categories and series used by shared OfficeIMO visual renderers.
/// </summary>
public sealed class OfficeChartData {
    /// <summary>
    /// Initializes a chart data snapshot.
    /// </summary>
    /// <param name="categories">Category labels, or numeric X values encoded as invariant strings for scatter charts.</param>
    /// <param name="series">Series values to render.</param>
    public OfficeChartData(IEnumerable<string> categories, IEnumerable<OfficeChartSeries> series) {
        if (categories == null) {
            throw new ArgumentNullException(nameof(categories));
        }

        if (series == null) {
            throw new ArgumentNullException(nameof(series));
        }

        Categories = new ReadOnlyCollection<string>(new List<string>(categories));
        Series = new ReadOnlyCollection<OfficeChartSeries>(new List<OfficeChartSeries>(series));
        if (Categories.Count == 0) {
            throw new ArgumentException("Chart data requires at least one category.", nameof(categories));
        }

        if (Series.Count == 0) {
            throw new ArgumentException("Chart data requires at least one series.", nameof(series));
        }
    }

    /// <summary>Category labels, or numeric X values encoded as invariant strings for scatter charts.</summary>
    public IReadOnlyList<string> Categories { get; }

    /// <summary>Series values aligned with chart categories.</summary>
    public IReadOnlyList<OfficeChartSeries> Series { get; }
}
