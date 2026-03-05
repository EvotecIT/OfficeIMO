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

    /// <summary>
    /// Describes X/Y series values for PowerPoint scatter charts.
    /// </summary>
    public sealed class PowerPointScatterChartData {
        /// <summary>
        /// Creates default scatter chart data used by parameterless chart creation.
        /// </summary>
        public static PowerPointScatterChartData CreateDefault() {
            return new PowerPointScatterChartData(new[] {
                new PowerPointScatterChartSeries("Series 1", new[] { 1d, 2d, 3d, 4d }, new[] { 2d, 4d, 3d, 5d }),
                new PowerPointScatterChartSeries("Series 2", new[] { 1d, 2d, 3d, 4d }, new[] { 1d, 3d, 2d, 4d })
            });
        }

        /// <summary>
        /// Initializes a new scatter chart data container with series.
        /// </summary>
        public PowerPointScatterChartData(IEnumerable<PowerPointScatterChartSeries> series) {
            if (series == null) {
                throw new ArgumentNullException(nameof(series));
            }

            Series = series.ToList();
            if (Series.Count == 0) {
                throw new ArgumentException("At least one series is required.", nameof(series));
            }
        }

        /// <summary>
        /// Series definitions for the scatter chart.
        /// </summary>
        public IReadOnlyList<PowerPointScatterChartSeries> Series { get; }

        /// <summary>
        /// Builds scatter chart data from a sequence of objects using selectors.
        /// </summary>
        public static PowerPointScatterChartData From<T>(IEnumerable<T> items, Func<T, double> xSelector,
            params PowerPointScatterChartSeriesDefinition<T>[] seriesDefinitions) {
            if (items == null) {
                throw new ArgumentNullException(nameof(items));
            }
            if (xSelector == null) {
                throw new ArgumentNullException(nameof(xSelector));
            }
            if (seriesDefinitions == null || seriesDefinitions.Length == 0) {
                throw new ArgumentException("At least one series definition is required.", nameof(seriesDefinitions));
            }

            var list = items.ToList();
            var xValues = list.Select(xSelector).ToList();
            var series = seriesDefinitions.Select(def =>
                new PowerPointScatterChartSeries(def.Name, xValues, list.Select(def.YSelector))).ToList();

            return new PowerPointScatterChartData(series);
        }
    }

    /// <summary>
    /// Represents a single scatter chart series with X/Y values.
    /// </summary>
    public sealed class PowerPointScatterChartSeries {
        /// <summary>
        /// Initializes a new scatter chart series with a name and X/Y values.
        /// </summary>
        public PowerPointScatterChartSeries(string name, IEnumerable<double> xValues, IEnumerable<double> yValues) {
            if (name == null) {
                throw new ArgumentNullException(nameof(name));
            }
            if (xValues == null) {
                throw new ArgumentNullException(nameof(xValues));
            }
            if (yValues == null) {
                throw new ArgumentNullException(nameof(yValues));
            }

            Name = name;
            XValues = xValues.ToList();
            YValues = yValues.ToList();

            if (XValues.Count == 0) {
                throw new ArgumentException("At least one X value is required.", nameof(xValues));
            }
            if (YValues.Count == 0) {
                throw new ArgumentException("At least one Y value is required.", nameof(yValues));
            }
            if (XValues.Count != YValues.Count) {
                throw new ArgumentException("X and Y value counts must match for each scatter series.");
            }
        }

        /// <summary>
        /// Series display name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Numeric X-axis values for the series.
        /// </summary>
        public IReadOnlyList<double> XValues { get; }

        /// <summary>
        /// Numeric Y-axis values for the series.
        /// </summary>
        public IReadOnlyList<double> YValues { get; }
    }

    /// <summary>
    /// Describes a scatter series for chart data generation from objects.
    /// </summary>
    public sealed class PowerPointScatterChartSeriesDefinition<T> {
        /// <summary>
        /// Initializes a scatter series definition.
        /// </summary>
        public PowerPointScatterChartSeriesDefinition(string name, Func<T, double> ySelector) {
            Name = name ?? throw new ArgumentNullException(nameof(name));
            YSelector = ySelector ?? throw new ArgumentNullException(nameof(ySelector));
        }

        /// <summary>
        /// Series name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Selector for Y-axis values.
        /// </summary>
        public Func<T, double> YSelector { get; }
    }
}
