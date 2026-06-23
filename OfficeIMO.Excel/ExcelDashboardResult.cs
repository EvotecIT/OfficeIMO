namespace OfficeIMO.Excel {
    /// <summary>
    /// Result of a dashboard build operation.
    /// </summary>
    public sealed class ExcelDashboardResult {
        internal ExcelDashboardResult(string tableRange, string? tableName, ExcelChart? chart) {
            TableRange = tableRange;
            TableName = tableName;
            Chart = chart;
        }

        /// <summary>A1 range occupied by the generated table.</summary>
        public string TableRange { get; }

        /// <summary>Name assigned to the generated table, when available.</summary>
        public string? TableName { get; }

        /// <summary>Generated chart, when requested.</summary>
        public ExcelChart? Chart { get; }
    }
}
