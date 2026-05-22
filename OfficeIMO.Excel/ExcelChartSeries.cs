using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Represents a chart series for Excel charts.
    /// </summary>
    public sealed class ExcelChartSeries {
        /// <summary>
        /// Creates a chart series with the specified name and values.
        /// </summary>
        public ExcelChartSeries(string name, IEnumerable<double> values, ExcelChartType? chartType = null, ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary)
            : this(name, (values ?? Array.Empty<double>()).ToList(), chartType, axisGroup, ownsValues: true) {
        }

        private ExcelChartSeries(string name, IReadOnlyList<double> values, ExcelChartType? chartType, ExcelChartAxisGroup axisGroup, bool ownsValues) {
            Name = name ?? string.Empty;
            Values = values ?? Array.Empty<double>();
            ChartType = chartType;
            AxisGroup = axisGroup;
        }

        internal static ExcelChartSeries CreateOwned(string name, IReadOnlyList<double> values, ExcelChartType? chartType = null, ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary)
            => new(name, values, chartType, axisGroup, ownsValues: true);

        /// <summary>
        /// Gets the series name.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the series values.
        /// </summary>
        public IReadOnlyList<double> Values { get; }

        /// <summary>
        /// Gets the optional chart type override for this series.
        /// </summary>
        public ExcelChartType? ChartType { get; }

        /// <summary>
        /// Gets the axis group for this series.
        /// </summary>
        public ExcelChartAxisGroup AxisGroup { get; }
    }
}
