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
        public ExcelChartData(IEnumerable<string> categories, IEnumerable<ExcelChartSeries> series) {
            Categories = (categories ?? Array.Empty<string>()).ToList();
            Series = (series ?? Array.Empty<ExcelChartSeries>()).ToList();

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

            var list = items as IList<T> ?? items.ToList();
            var categories = new List<string>(list.Count);
            for (int i = 0; i < list.Count; i++) {
                categories.Add(categorySelector(list[i]));
            }

            var series = new List<ExcelChartSeries>(seriesDefinitions.Length);
            foreach (var def in seriesDefinitions) {
                if (def == null) {
                    throw new ArgumentNullException(nameof(seriesDefinitions));
                }
                var values = new double[list.Count];
                for (int i = 0; i < list.Count; i++) {
                    values[i] = def.ValueSelector(list[i]);
                }
                series.Add(new ExcelChartSeries(def.Name, values, def.ChartType, def.AxisGroup));
            }

            return new ExcelChartData(categories, series);
        }
    }
}
