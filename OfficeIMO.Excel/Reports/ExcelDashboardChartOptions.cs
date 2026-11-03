namespace OfficeIMO.Excel.Reports {
    /// <summary>
    /// Common dashboard chart workflows.
    /// </summary>
    public enum ExcelDashboardChartPreset {
        /// <summary>Clustered column chart for comparing categories.</summary>
        Comparison,
        /// <summary>Line chart for trends over time or ordered categories.</summary>
        Trend,
        /// <summary>Pie chart for contribution/share views.</summary>
        Contribution,
        /// <summary>Compact clustered bar chart for dashboards with limited space.</summary>
        CompactComparison
    }

    /// <summary>
    /// Options for adding a dashboard-ready chart.
    /// </summary>
    public sealed class ExcelDashboardChartOptions {
        /// <summary>Dashboard chart preset.</summary>
        public ExcelDashboardChartPreset Preset { get; set; } = ExcelDashboardChartPreset.Comparison;

        /// <summary>A1 range containing chart data.</summary>
        public string? Range { get; set; }

        /// <summary>Table name containing chart data.</summary>
        public string? TableName { get; set; }

        /// <summary>Top-left row for the chart.</summary>
        public int Row { get; set; } = 1;

        /// <summary>Top-left column for the chart.</summary>
        public int Column { get; set; } = 1;

        /// <summary>Optional chart type override.</summary>
        public OfficeChartType? ChartType { get; set; }

        /// <summary>Optional chart title.</summary>
        public string? Title { get; set; }

        /// <summary>Whether the source range has headers.</summary>
        public bool HasHeaders { get; set; } = true;

        /// <summary>Include cached data in the chart for portability.</summary>
        public bool IncludeCachedData { get; set; } = true;

        /// <summary>Optional width physical override.</summary>
        public OfficeLength? Width { get; set; }

        /// <summary>Optional height physical override.</summary>
        public OfficeLength? Height { get; set; }

        /// <summary>Optional chart style id override.</summary>
        public int? StyleId { get; set; }

        /// <summary>Optional chart color style id override.</summary>
        public int? ColorStyleId { get; set; }
    }
}
