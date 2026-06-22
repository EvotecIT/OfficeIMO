namespace OfficeIMO.Excel {
    /// <summary>
    /// Options for building a dashboard table and chart from tabular data.
    /// </summary>
    public sealed class ExcelDashboardOptions {
        /// <summary>Optional dashboard title.</summary>
        public string? Title { get; set; }

        /// <summary>Optional dashboard subtitle.</summary>
        public string? Subtitle { get; set; }

        /// <summary>Top-left row for the generated table.</summary>
        public int TableRow { get; set; } = 3;

        /// <summary>Top-left column for the generated table.</summary>
        public int TableColumn { get; set; } = 1;

        /// <summary>Optional table name.</summary>
        public string? TableName { get; set; }

        /// <summary>Built-in table style.</summary>
        public TableStyle TableStyle { get; set; } = TableStyle.TableStyleMedium9;

        /// <summary>Whether the table should include AutoFilter dropdowns.</summary>
        public bool IncludeAutoFilter { get; set; } = true;

        /// <summary>Whether table columns should be auto-fitted after insertion.</summary>
        public bool AutoFit { get; set; } = true;

        /// <summary>Whether to add a dashboard chart beside the table.</summary>
        public bool AddChart { get; set; } = true;

        /// <summary>Dashboard chart preset.</summary>
        public ExcelDashboardChartPreset ChartPreset { get; set; } = ExcelDashboardChartPreset.Comparison;

        /// <summary>Optional chart title. Defaults to <see cref="Title"/> when omitted.</summary>
        public string? ChartTitle { get; set; }

        /// <summary>Optional top-left chart row. Defaults to <see cref="TableRow"/>.</summary>
        public int? ChartRow { get; set; }

        /// <summary>Optional top-left chart column. Defaults to two columns after the table.</summary>
        public int? ChartColumn { get; set; }
    }
}
