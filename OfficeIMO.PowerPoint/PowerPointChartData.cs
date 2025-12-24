using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.PowerPoint {
    /// <summary>
    /// Describes chart categories and series values for PowerPoint charts.
    /// </summary>
    public sealed class PowerPointChartData {
        /// <summary>
        /// Creates default chart data used by parameterless chart creation.
        /// </summary>
        public static PowerPointChartData CreateDefault() {
            return new PowerPointChartData(
                new[] { "Category 1", "Category 2", "Category 3", "Category 4" },
                new[] {
                    new PowerPointChartSeries("Series 1", new[] { 4d, 2d, 3d, 5d }),
                    new PowerPointChartSeries("Series 2", new[] { 2d, 4d, 2d, 3d }),
                    new PowerPointChartSeries("Series 3", new[] { 1d, 3d, 2d, 4d })
                });
        }

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

        /// <summary>
        /// Builds chart data from a sequence of objects using selectors.
        /// </summary>
        public static PowerPointChartData From<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            params PowerPointChartSeriesDefinition<T>[] seriesDefinitions) {
            if (items == null) {
                throw new System.ArgumentNullException(nameof(items));
            }
            if (categorySelector == null) {
                throw new System.ArgumentNullException(nameof(categorySelector));
            }
            if (seriesDefinitions == null || seriesDefinitions.Length == 0) {
                throw new System.ArgumentException("At least one series definition is required.", nameof(seriesDefinitions));
            }

            var list = items.ToList();
            var categories = list.Select(categorySelector).ToList();
            var series = seriesDefinitions.Select(def =>
                new PowerPointChartSeries(def.Name, list.Select(def.ValueSelector))).ToList();

            return new PowerPointChartData(categories, series);
        }
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

    /// <summary>
    /// Describes a series for chart data generation from objects.
    /// </summary>
    public sealed class PowerPointChartSeriesDefinition<T> {
        /// <summary>
        /// Initializes a series definition.
        /// </summary>
        public PowerPointChartSeriesDefinition(string name, Func<T, double> valueSelector) {
            Name = name ?? throw new System.ArgumentNullException(nameof(name));
            ValueSelector = valueSelector ?? throw new System.ArgumentNullException(nameof(valueSelector));
        }

        /// <summary>
        /// Series name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Selector for series values.
        /// </summary>
        public Func<T, double> ValueSelector { get; }
    }
}
