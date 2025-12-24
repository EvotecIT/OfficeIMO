using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Describes chart categories and series values for PowerPoint charts.
    /// </summary>
    public sealed class PowerPointChartData {
        /// <summary>
        /// Initializes a new chart data container with categories and series.
        /// </summary>
        public PowerPointChartData(IEnumerable<string> categories, IEnumerable<PowerPointChartSeries> series) {
            if (categories == null) {
                throw new System.ArgumentNullException(nameof(categories));
            }

            if (series == null) {
                throw new System.ArgumentNullException(nameof(series));
            }

            Categories = categories.ToList();
            Series = series.ToList();

            if (Categories.Count == 0) {
                throw new System.ArgumentException("At least one category is required.", nameof(categories));
            }

            if (Series.Count == 0) {
                throw new System.ArgumentException("At least one series is required.", nameof(series));
            }

            int expected = Categories.Count;
            foreach (PowerPointChartSeries item in Series) {
                if (item.Values.Count != expected) {
                    throw new System.ArgumentException(
                        "Each series must have the same number of values as there are categories.",
                        nameof(series));
                }
            }
        }

        /// <summary>
        /// Category labels for the chart.
        /// </summary>
        public IReadOnlyList<string> Categories { get; }

        /// <summary>
        /// Series definitions for the chart.
        /// </summary>
        public IReadOnlyList<PowerPointChartSeries> Series { get; }
    }

    /// <summary>
    /// Represents a single chart series.
    /// </summary>
    public sealed class PowerPointChartSeries {
        /// <summary>
        /// Initializes a new chart series with a name and values.
        /// </summary>
        public PowerPointChartSeries(string name, IEnumerable<double> values) {
            if (name == null) {
                throw new System.ArgumentNullException(nameof(name));
            }

            if (values == null) {
                throw new System.ArgumentNullException(nameof(values));
            }

            Name = name;
            Values = values.ToList();
        }

        /// <summary>
        /// Series display name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Series values aligned with chart categories.
        /// </summary>
        public IReadOnlyList<double> Values { get; }
    }
}
