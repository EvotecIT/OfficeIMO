using System;

namespace OfficeIMO.Excel {
    /// <summary>
    /// Defines a series mapping for building chart data from objects.
    /// </summary>
    public sealed class ExcelChartSeriesDefinition<T> {
        /// <summary>
        /// Creates a series definition that maps values from the source items.
        /// </summary>
        public ExcelChartSeriesDefinition(string name, Func<T, double> valueSelector, ExcelChartType? chartType = null, ExcelChartAxisGroup axisGroup = ExcelChartAxisGroup.Primary) {
            Name = name ?? string.Empty;
            ValueSelector = valueSelector ?? throw new ArgumentNullException(nameof(valueSelector));
            ChartType = chartType;
            AxisGroup = axisGroup;
        }

        /// <summary>
        /// Gets the display name of the series.
        /// </summary>
        public string Name { get; }

        /// <summary>
        /// Gets the value selector for the series.
        /// </summary>
        public Func<T, double> ValueSelector { get; }

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
