namespace OfficeIMO.Excel {
    /// <summary>
    /// Supported Excel chart types for OfficeIMO.Excel chart helpers.
    /// </summary>
    public enum ExcelChartType {
        /// <summary>Clustered column (vertical bars).</summary>
        ColumnClustered,
        /// <summary>Stacked column (vertical bars).</summary>
        ColumnStacked,
        /// <summary>Clustered bar (horizontal bars).</summary>
        BarClustered,
        /// <summary>Stacked bar (horizontal bars).</summary>
        BarStacked,
        /// <summary>Line chart.</summary>
        Line,
        /// <summary>Area chart.</summary>
        Area,
        /// <summary>Pie chart.</summary>
        Pie,
        /// <summary>Doughnut chart.</summary>
        Doughnut,
        /// <summary>Scatter (XY) chart.</summary>
        Scatter,
        /// <summary>Bubble chart.</summary>
        Bubble
    }
}
