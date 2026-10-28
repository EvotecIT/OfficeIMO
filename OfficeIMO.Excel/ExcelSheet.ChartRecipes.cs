using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        /// <summary>
        /// Adds a line chart with dashboard-friendly defaults for time-series revenue or volume trends.
        /// </summary>
        public ExcelChart AddRevenueTrendChart(string dataRangeA1, int row, int column, string title = "Revenue Trend",
            OfficeLength? width = null, OfficeLength? height = null, bool hasHeaders = true, bool includeCachedData = true) {
            width ??= OfficeLength.Pixels(720);
            height ??= OfficeLength.Pixels(320);
            return AddChartFromRange(dataRangeA1, row, column, width.Value, height.Value, OfficeChartType.Line, hasHeaders, title, includeCachedData);
        }

        /// <summary>
        /// Adds a doughnut chart with dashboard-friendly defaults for status, category, or allocation breakdowns.
        /// </summary>
        public ExcelChart AddStatusBreakdownChart(string dataRangeA1, int row, int column, string title = "Status Breakdown",
            OfficeLength? width = null, OfficeLength? height = null, bool hasHeaders = true, bool includeCachedData = true) {
            width ??= OfficeLength.Pixels(520);
            height ??= OfficeLength.Pixels(320);
            return AddChartFromRange(dataRangeA1, row, column, width.Value, height.Value, OfficeChartType.Doughnut, hasHeaders, title, includeCachedData);
        }

        /// <summary>
        /// Adds a horizontal bar chart with dashboard-friendly defaults for top-N rankings.
        /// </summary>
        public ExcelChart AddTopNBarChart(string dataRangeA1, int row, int column, string title = "Top Items",
            OfficeLength? width = null, OfficeLength? height = null, bool hasHeaders = true, bool includeCachedData = true) {
            width ??= OfficeLength.Pixels(640);
            height ??= OfficeLength.Pixels(360);
            return AddChartFromRange(dataRangeA1, row, column, width.Value, height.Value, OfficeChartType.BarClustered, hasHeaders, title, includeCachedData);
        }

        /// <summary>
        /// Adds a clustered column chart with dashboard-friendly defaults for variance comparisons.
        /// </summary>
        public ExcelChart AddVarianceColumnChart(string dataRangeA1, int row, int column, string title = "Variance",
            OfficeLength? width = null, OfficeLength? height = null, bool hasHeaders = true, bool includeCachedData = true) {
            width ??= OfficeLength.Pixels(640);
            height ??= OfficeLength.Pixels(360);
            return AddChartFromRange(dataRangeA1, row, column, width.Value, height.Value, OfficeChartType.ColumnClustered, hasHeaders, title, includeCachedData);
        }

        /// <summary>
        /// Adds a compact column chart with value callouts and dashboard-friendly defaults for KPI scorecards.
        /// </summary>
        public ExcelChart AddKpiScorecardChart(string dataRangeA1, int row, int column, string title = "KPI Scorecard",
            OfficeLength? width = null, OfficeLength? height = null, bool hasHeaders = true, bool includeCachedData = true) {
            width ??= OfficeLength.Pixels(520);
            height ??= OfficeLength.Pixels(300);
            ExcelChart chart = AddChartFromRange(dataRangeA1, row, column, width.Value, height.Value, OfficeChartType.ColumnClustered, hasHeaders, title, includeCachedData);
            return ApplyKpiScorecardDefaults(chart);
        }

        /// <summary>
        /// Adds a doughnut chart with category/percent labels for contribution and mix analysis.
        /// </summary>
        public ExcelChart AddContributionChart(string dataRangeA1, int row, int column, string title = "Contribution",
            OfficeLength? width = null, OfficeLength? height = null, bool hasHeaders = true, bool includeCachedData = true) {
            width ??= OfficeLength.Pixels(520);
            height ??= OfficeLength.Pixels(320);
            ExcelChart chart = AddChartFromRange(dataRangeA1, row, column, width.Value, height.Value, OfficeChartType.Doughnut, hasHeaders, title, includeCachedData);
            return ApplyContributionChartDefaults(chart);
        }

        /// <summary>
        /// Adds a waterfall-style stacked column chart for variance bridges prepared with helper series.
        /// </summary>
        public ExcelChart AddVarianceWaterfallChart(string dataRangeA1, int row, int column, string title = "Variance Bridge",
            OfficeLength? width = null, OfficeLength? height = null, bool hasHeaders = true, bool includeCachedData = true) {
            width ??= OfficeLength.Pixels(720);
            height ??= OfficeLength.Pixels(360);
            ExcelChart chart = AddChartFromRange(dataRangeA1, row, column, width.Value, height.Value, OfficeChartType.ColumnStacked, hasHeaders, title, includeCachedData);
            return ApplyVarianceWaterfallDefaults(chart);
        }

        internal static ExcelChart ApplyKpiScorecardDefaults(ExcelChart chart) {
            return chart.HideLegend()
                .SetTitleTextStyle(fontSize: OfficeLength.Points(12), bold: true, color: OfficeColor.Parse("1F2937"))
                .SetSeriesDataLabels(0, showValue: true, position: OfficeChartDataLabelPosition.OutsideEnd, numberFormat: "#,##0")
                .SetSeriesDataLabelTextStyle(0, fontSize: OfficeLength.Points(10), bold: true, color: OfficeColor.Parse("1F2937"))
                .SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: OfficeColor.Parse("E5E7EB"), lineWidth: OfficeLength.Points(0.5))
                .SetValueAxisNumberFormat("#,##0", sourceLinked: false);
        }

        internal static ExcelChart ApplyContributionChartDefaults(ExcelChart chart) {
            return chart.SetTitleTextStyle(fontSize: OfficeLength.Points(12), bold: true, color: OfficeColor.Parse("1F2937"))
                .SetLegend(OfficeChartLegendPosition.Right)
                .SetSeriesDataLabels(0, showValue: false, showCategoryName: true, showPercent: true, position: OfficeChartDataLabelPosition.BestFit, numberFormat: "0%")
                .SetSeriesDataLabelSeparator(0, "\n")
                .SetSeriesDataLabelTextStyle(0, fontSize: OfficeLength.Points(9), color: OfficeColor.Parse("374151"));
        }

        internal static ExcelChart ApplyVarianceWaterfallDefaults(ExcelChart chart) {
            return chart.SetTitleTextStyle(fontSize: OfficeLength.Points(12), bold: true, color: OfficeColor.Parse("1F2937"))
                .SetLegend(OfficeChartLegendPosition.Bottom)
                .SetSeriesDataLabels(0, showValue: true, position: OfficeChartDataLabelPosition.OutsideEnd, numberFormat: "#,##0")
                .SetSeriesDataLabelTextStyle(0, fontSize: OfficeLength.Points(9), color: OfficeColor.Parse("374151"))
                .SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: OfficeColor.Parse("E5E7EB"), lineWidth: OfficeLength.Points(0.5))
                .SetValueAxisNumberFormat("#,##0", sourceLinked: false);
        }
    }
}
