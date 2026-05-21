namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Adds a line chart with dashboard-friendly defaults for time-series revenue or volume trends.
        /// </summary>
        public ExcelChart AddRevenueTrendChart(string dataRangeA1, int row, int column, string title = "Revenue Trend",
            int widthPixels = 720, int heightPixels = 320, bool hasHeaders = true, bool includeCachedData = true) {
            return AddChartFromRange(dataRangeA1, row, column, widthPixels, heightPixels, ExcelChartType.Line, hasHeaders, title, includeCachedData);
        }

        /// <summary>
        /// Adds a doughnut chart with dashboard-friendly defaults for status, category, or allocation breakdowns.
        /// </summary>
        public ExcelChart AddStatusBreakdownChart(string dataRangeA1, int row, int column, string title = "Status Breakdown",
            int widthPixels = 520, int heightPixels = 320, bool hasHeaders = true, bool includeCachedData = true) {
            return AddChartFromRange(dataRangeA1, row, column, widthPixels, heightPixels, ExcelChartType.Doughnut, hasHeaders, title, includeCachedData);
        }

        /// <summary>
        /// Adds a horizontal bar chart with dashboard-friendly defaults for top-N rankings.
        /// </summary>
        public ExcelChart AddTopNBarChart(string dataRangeA1, int row, int column, string title = "Top Items",
            int widthPixels = 640, int heightPixels = 360, bool hasHeaders = true, bool includeCachedData = true) {
            return AddChartFromRange(dataRangeA1, row, column, widthPixels, heightPixels, ExcelChartType.BarClustered, hasHeaders, title, includeCachedData);
        }

        /// <summary>
        /// Adds a clustered column chart with dashboard-friendly defaults for variance comparisons.
        /// </summary>
        public ExcelChart AddVarianceColumnChart(string dataRangeA1, int row, int column, string title = "Variance",
            int widthPixels = 640, int heightPixels = 360, bool hasHeaders = true, bool includeCachedData = true) {
            return AddChartFromRange(dataRangeA1, row, column, widthPixels, heightPixels, ExcelChartType.ColumnClustered, hasHeaders, title, includeCachedData);
        }
    }
}
