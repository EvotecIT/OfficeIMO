using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Excel {
    /// <summary>
    /// In-memory chart data for Excel charts.
    /// </summary>
    public sealed class ExcelChartData {
        /// <summary>
        /// Creates chart data from the provided categories and series.
        /// </summary>
        public ExcelChartData(IEnumerable<string> categories, IEnumerable<ExcelChartSeries> series)
            : this((categories ?? Array.Empty<string>()).ToList(), (series ?? Array.Empty<ExcelChartSeries>()).ToList(), ownsData: true) {
        }

        private ExcelChartData(IReadOnlyList<string> categories, IReadOnlyList<ExcelChartSeries> series, bool ownsData) {
            Categories = categories ?? Array.Empty<string>();
            Series = series ?? Array.Empty<ExcelChartSeries>();
            if (Series.Count == 0) {
                throw new ArgumentException("At least one series is required.", nameof(series));
            }

            int count = Categories.Count;
            foreach (var item in Series) {
                if (item.Values.Count != count) {
                    throw new ArgumentException("Each series must match the categories count.", nameof(series));
                }
            }
        }

        /// <summary>
        /// Gets the category labels for the chart.
        /// </summary>
        public IReadOnlyList<string> Categories { get; }

        /// <summary>
        /// Gets the chart series collection.
        /// </summary>
        public IReadOnlyList<ExcelChartSeries> Series { get; }

        /// <summary>
        /// Creates a sample chart dataset.
        /// </summary>
        public static ExcelChartData CreateDefault() {
            return new ExcelChartData(
                new[] { "Category 1", "Category 2", "Category 3", "Category 4" },
                new[] {
                    new ExcelChartSeries("Series 1", new[] { 4d, 2d, 3d, 5d }),
                    new ExcelChartSeries("Series 2", new[] { 2d, 4d, 2d, 3d }),
                    new ExcelChartSeries("Series 3", new[] { 1d, 3d, 2d, 4d })
                });
        }

        /// <summary>
        /// Builds chart data from a sequence of items and series definitions.
        /// </summary>
        public static ExcelChartData From<T>(IEnumerable<T> items, Func<T, string> categorySelector,
            params ExcelChartSeriesDefinition<T>[] seriesDefinitions) {
            if (items == null) throw new ArgumentNullException(nameof(items));
            if (categorySelector == null) throw new ArgumentNullException(nameof(categorySelector));
            if (seriesDefinitions == null || seriesDefinitions.Length == 0) {
                throw new ArgumentException("At least one series definition is required.", nameof(seriesDefinitions));
            }

            ValidateSeriesDefinitions(seriesDefinitions);
            if (items is IReadOnlyList<T> list) {
                return FromIndexedList(list, categorySelector, seriesDefinitions);
            }

            return FromEnumerable(items, categorySelector, seriesDefinitions);
        }

        private static void ValidateSeriesDefinitions<T>(IReadOnlyList<ExcelChartSeriesDefinition<T>> seriesDefinitions) {
            for (int i = 0; i < seriesDefinitions.Count; i++) {
                if (seriesDefinitions[i] == null) {
                    throw new ArgumentNullException(nameof(seriesDefinitions));
                }
            }
        }

        private static ExcelChartData CreateOwned(IReadOnlyList<string> categories, IReadOnlyList<ExcelChartSeries> series)
            => new(categories, series, ownsData: true);

        private static ExcelChartData FromIndexedList<T>(
            IReadOnlyList<T> list,
            Func<T, string> categorySelector,
            IReadOnlyList<ExcelChartSeriesDefinition<T>> seriesDefinitions) {
            var categories = new List<string>(list.Count);
            for (int i = 0; i < list.Count; i++) {
                categories.Add(categorySelector(list[i]));
            }

            var series = new List<ExcelChartSeries>(seriesDefinitions.Count);
            for (int seriesIndex = 0; seriesIndex < seriesDefinitions.Count; seriesIndex++) {
                var def = seriesDefinitions[seriesIndex];
                var values = new double[list.Count];
                for (int i = 0; i < list.Count; i++) {
                    values[i] = def.ValueSelector(list[i]);
                }

                series.Add(ExcelChartSeries.CreateOwned(def.Name, values, def.ChartType, def.AxisGroup));
            }

            return CreateOwned(categories, series);
        }

        private static ExcelChartData FromEnumerable<T>(
            IEnumerable<T> items,
            Func<T, string> categorySelector,
            IReadOnlyList<ExcelChartSeriesDefinition<T>> seriesDefinitions) {
            int capacity = items is IReadOnlyCollection<T> readOnlyCollection
                ? readOnlyCollection.Count
                : items is ICollection<T> collection ? collection.Count : 0;
            var categories = capacity > 0 ? new List<string>(capacity) : new List<string>();
            var seriesValues = new List<double>[seriesDefinitions.Count];
            for (int i = 0; i < seriesValues.Length; i++) {
                seriesValues[i] = capacity > 0 ? new List<double>(capacity) : new List<double>();
            }

            foreach (T item in items) {
                categories.Add(categorySelector(item));
                for (int i = 0; i < seriesDefinitions.Count; i++) {
                    seriesValues[i].Add(seriesDefinitions[i].ValueSelector(item));
                }
            }

            var series = new List<ExcelChartSeries>(seriesDefinitions.Count);
            for (int i = 0; i < seriesDefinitions.Count; i++) {
                var def = seriesDefinitions[i];
                series.Add(ExcelChartSeries.CreateOwned(def.Name, seriesValues[i], def.ChartType, def.AxisGroup));
            }

            return CreateOwned(categories, series);
        }
    }
}
