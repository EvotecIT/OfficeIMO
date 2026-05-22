using C = DocumentFormat.OpenXml.Drawing.Charts;

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

        /// <summary>
        /// Adds a compact column chart with value callouts and dashboard-friendly defaults for KPI scorecards.
        /// </summary>
        public ExcelChart AddKpiScorecardChart(string dataRangeA1, int row, int column, string title = "KPI Scorecard",
            int widthPixels = 520, int heightPixels = 300, bool hasHeaders = true, bool includeCachedData = true) {
            ExcelChart chart = AddChartFromRange(dataRangeA1, row, column, widthPixels, heightPixels, ExcelChartType.ColumnClustered, hasHeaders, title, includeCachedData);
            return ApplyKpiScorecardDefaults(chart);
        }

        /// <summary>
        /// Adds a doughnut chart with category/percent labels for contribution and mix analysis.
        /// </summary>
        public ExcelChart AddContributionChart(string dataRangeA1, int row, int column, string title = "Contribution",
            int widthPixels = 520, int heightPixels = 320, bool hasHeaders = true, bool includeCachedData = true) {
            ExcelChart chart = AddChartFromRange(dataRangeA1, row, column, widthPixels, heightPixels, ExcelChartType.Doughnut, hasHeaders, title, includeCachedData);
            return ApplyContributionChartDefaults(chart);
        }

        /// <summary>
        /// Adds a waterfall-style stacked column chart for variance bridges prepared with helper series.
        /// </summary>
        public ExcelChart AddVarianceWaterfallChart(string dataRangeA1, int row, int column, string title = "Variance Bridge",
            int widthPixels = 720, int heightPixels = 360, bool hasHeaders = true, bool includeCachedData = true) {
            ExcelChart chart = AddChartFromRange(dataRangeA1, row, column, widthPixels, heightPixels, ExcelChartType.ColumnStacked, hasHeaders, title, includeCachedData);
            return ApplyVarianceWaterfallDefaults(chart);
        }

        internal static ExcelChart ApplyKpiScorecardDefaults(ExcelChart chart) {
            return chart.HideLegend()
                .SetTitleTextStyle(fontSizePoints: 12, bold: true, color: "1F2937")
                .SetSeriesDataLabels(0, showValue: true, position: C.DataLabelPositionValues.OutsideEnd, numberFormat: "#,##0")
                .SetSeriesDataLabelTextStyle(0, fontSizePoints: 10, bold: true, color: "1F2937")
                .SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: "E5E7EB", lineWidthPoints: 0.5)
                .SetValueAxisNumberFormat("#,##0", sourceLinked: false);
        }

        internal static ExcelChart ApplyContributionChartDefaults(ExcelChart chart) {
            return chart.SetTitleTextStyle(fontSizePoints: 12, bold: true, color: "1F2937")
                .SetLegend(C.LegendPositionValues.Right)
                .SetSeriesDataLabels(0, showValue: false, showCategoryName: true, showPercent: true, position: C.DataLabelPositionValues.BestFit, numberFormat: "0%")
                .SetSeriesDataLabelSeparator(0, "\n")
                .SetSeriesDataLabelTextStyle(0, fontSizePoints: 9, color: "374151");
        }

        internal static ExcelChart ApplyVarianceWaterfallDefaults(ExcelChart chart) {
            return chart.SetTitleTextStyle(fontSizePoints: 12, bold: true, color: "1F2937")
                .SetLegend(C.LegendPositionValues.Bottom)
                .SetSeriesDataLabels(0, showValue: true, position: C.DataLabelPositionValues.OutsideEnd, numberFormat: "#,##0")
                .SetSeriesDataLabelTextStyle(0, fontSizePoints: 9, color: "374151")
                .SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: "E5E7EB", lineWidthPoints: 0.5)
                .SetValueAxisNumberFormat("#,##0", sourceLinked: false);
        }
    }
}
