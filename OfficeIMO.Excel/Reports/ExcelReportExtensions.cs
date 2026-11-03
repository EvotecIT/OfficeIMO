using System.Data;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Excel.Reports {
    /// <summary>
    /// Business report presets for worksheets: KPI, trend, variance and contribution charts, dashboards and subtotal
    /// summaries. Each preset is built on the general chart, table and outline APIs of <see cref="ExcelSheet"/>.
    /// </summary>
    public static class ExcelReportExtensions {
        /// <summary>
        /// Adds a line chart with dashboard-friendly defaults for time-series revenue or volume trends.
        /// </summary>
        public static ExcelChart AddRevenueTrendChart(this ExcelSheet sheet, string dataRangeA1, int row, int column, string title = "Revenue Trend",
            OfficeLength? width = null, OfficeLength? height = null, bool hasHeaders = true, bool includeCachedData = true) {
            width ??= OfficeLength.Pixels(720);
            height ??= OfficeLength.Pixels(320);
            return sheet.AddChart(OfficeChartType.Line, dataRangeA1, new OfficeIMO.Excel.ExcelChartOptions { Row = row, Column = column, Width = width.Value, Height = height.Value, HasHeaders = hasHeaders, Title = title, IncludeCachedData = includeCachedData });
        }

        /// <summary>
        /// Adds a doughnut chart with dashboard-friendly defaults for status, category, or allocation breakdowns.
        /// </summary>
        public static ExcelChart AddStatusBreakdownChart(this ExcelSheet sheet, string dataRangeA1, int row, int column, string title = "Status Breakdown",
            OfficeLength? width = null, OfficeLength? height = null, bool hasHeaders = true, bool includeCachedData = true) {
            width ??= OfficeLength.Pixels(520);
            height ??= OfficeLength.Pixels(320);
            return sheet.AddChart(OfficeChartType.Doughnut, dataRangeA1, new OfficeIMO.Excel.ExcelChartOptions { Row = row, Column = column, Width = width.Value, Height = height.Value, HasHeaders = hasHeaders, Title = title, IncludeCachedData = includeCachedData });
        }

        /// <summary>
        /// Adds a horizontal bar chart with dashboard-friendly defaults for top-N rankings.
        /// </summary>
        public static ExcelChart AddTopNBarChart(this ExcelSheet sheet, string dataRangeA1, int row, int column, string title = "Top Items",
            OfficeLength? width = null, OfficeLength? height = null, bool hasHeaders = true, bool includeCachedData = true) {
            width ??= OfficeLength.Pixels(640);
            height ??= OfficeLength.Pixels(360);
            return sheet.AddChart(OfficeChartType.BarClustered, dataRangeA1, new OfficeIMO.Excel.ExcelChartOptions { Row = row, Column = column, Width = width.Value, Height = height.Value, HasHeaders = hasHeaders, Title = title, IncludeCachedData = includeCachedData });
        }

        /// <summary>
        /// Adds a clustered column chart with dashboard-friendly defaults for variance comparisons.
        /// </summary>
        public static ExcelChart AddVarianceColumnChart(this ExcelSheet sheet, string dataRangeA1, int row, int column, string title = "Variance",
            OfficeLength? width = null, OfficeLength? height = null, bool hasHeaders = true, bool includeCachedData = true) {
            width ??= OfficeLength.Pixels(640);
            height ??= OfficeLength.Pixels(360);
            return sheet.AddChart(OfficeChartType.ColumnClustered, dataRangeA1, new OfficeIMO.Excel.ExcelChartOptions { Row = row, Column = column, Width = width.Value, Height = height.Value, HasHeaders = hasHeaders, Title = title, IncludeCachedData = includeCachedData });
        }

        /// <summary>
        /// Adds a compact column chart with value callouts and dashboard-friendly defaults for KPI scorecards.
        /// </summary>
        public static ExcelChart AddKpiScorecardChart(this ExcelSheet sheet, string dataRangeA1, int row, int column, string title = "KPI Scorecard",
            OfficeLength? width = null, OfficeLength? height = null, bool hasHeaders = true, bool includeCachedData = true) {
            width ??= OfficeLength.Pixels(520);
            height ??= OfficeLength.Pixels(300);
            ExcelChart chart = sheet.AddChart(OfficeChartType.ColumnClustered, dataRangeA1, new OfficeIMO.Excel.ExcelChartOptions { Row = row, Column = column, Width = width.Value, Height = height.Value, HasHeaders = hasHeaders, Title = title, IncludeCachedData = includeCachedData });
            return ApplyKpiScorecardDefaults(chart);
        }

        /// <summary>
        /// Adds a doughnut chart with category/percent labels for contribution and mix analysis.
        /// </summary>
        public static ExcelChart AddContributionChart(this ExcelSheet sheet, string dataRangeA1, int row, int column, string title = "Contribution",
            OfficeLength? width = null, OfficeLength? height = null, bool hasHeaders = true, bool includeCachedData = true) {
            width ??= OfficeLength.Pixels(520);
            height ??= OfficeLength.Pixels(320);
            ExcelChart chart = sheet.AddChart(OfficeChartType.Doughnut, dataRangeA1, new OfficeIMO.Excel.ExcelChartOptions { Row = row, Column = column, Width = width.Value, Height = height.Value, HasHeaders = hasHeaders, Title = title, IncludeCachedData = includeCachedData });
            return ApplyContributionChartDefaults(chart);
        }

        /// <summary>
        /// Adds a waterfall-style stacked column chart for variance bridges prepared with helper series.
        /// </summary>
        public static ExcelChart AddVarianceWaterfallChart(this ExcelSheet sheet, string dataRangeA1, int row, int column, string title = "Variance Bridge",
            OfficeLength? width = null, OfficeLength? height = null, bool hasHeaders = true, bool includeCachedData = true) {
            width ??= OfficeLength.Pixels(720);
            height ??= OfficeLength.Pixels(360);
            ExcelChart chart = sheet.AddChart(OfficeChartType.ColumnStacked, dataRangeA1, new OfficeIMO.Excel.ExcelChartOptions { Row = row, Column = column, Width = width.Value, Height = height.Value, HasHeaders = hasHeaders, Title = title, IncludeCachedData = includeCachedData });
            return ApplyVarianceWaterfallDefaults(chart);
        }

        private static ExcelChart ApplyKpiScorecardDefaults(ExcelChart chart) {
            return chart.HideLegend()
                .SetTitleTextStyle(fontSize: OfficeLength.Points(12), bold: true, color: OfficeColor.Parse("1F2937"))
                .SetSeriesDataLabels(0, showValue: true, position: OfficeChartDataLabelPosition.OutsideEnd, numberFormat: "#,##0")
                .SetSeriesDataLabelTextStyle(0, fontSize: OfficeLength.Points(10), bold: true, color: OfficeColor.Parse("1F2937"))
                .SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: OfficeColor.Parse("E5E7EB"), lineWidth: OfficeLength.Points(0.5))
                .SetValueAxisNumberFormat("#,##0", sourceLinked: false);
        }

        private static ExcelChart ApplyContributionChartDefaults(ExcelChart chart) {
            return chart.SetTitleTextStyle(fontSize: OfficeLength.Points(12), bold: true, color: OfficeColor.Parse("1F2937"))
                .SetLegend(OfficeChartLegendPosition.Right)
                .SetSeriesDataLabels(0, showValue: false, showCategoryName: true, showPercent: true, position: OfficeChartDataLabelPosition.BestFit, numberFormat: "0%")
                .SetSeriesDataLabelSeparator(0, "\n")
                .SetSeriesDataLabelTextStyle(0, fontSize: OfficeLength.Points(9), color: OfficeColor.Parse("374151"));
        }

        private static ExcelChart ApplyVarianceWaterfallDefaults(ExcelChart chart) {
            return chart.SetTitleTextStyle(fontSize: OfficeLength.Points(12), bold: true, color: OfficeColor.Parse("1F2937"))
                .SetLegend(OfficeChartLegendPosition.Bottom)
                .SetSeriesDataLabels(0, showValue: true, position: OfficeChartDataLabelPosition.OutsideEnd, numberFormat: "#,##0")
                .SetSeriesDataLabelTextStyle(0, fontSize: OfficeLength.Points(9), color: OfficeColor.Parse("374151"))
                .SetValueAxisGridlines(showMajor: true, showMinor: false, lineColor: OfficeColor.Parse("E5E7EB"), lineWidth: OfficeLength.Points(0.5))
                .SetValueAxisNumberFormat("#,##0", sourceLinked: false);
        }

        /// <summary>Applies a dashboard preset's style, legend and data labels to an existing chart.</summary>
        /// <remarks>The chart's data, family and placement are kept. Style overrides use the same values as
        /// <see cref="AddDashboardChart"/>.</remarks>
        /// <param name="chart">Chart to restyle.</param>
        /// <param name="preset">The dashboard appearance preset.</param>
        /// <param name="styleId">Optional chart style override.</param>
        /// <param name="colorStyleId">Optional chart color-style override.</param>
        /// <returns>The same chart for further editing.</returns>
        public static ExcelChart ApplyDashboardPreset(this ExcelChart chart, ExcelDashboardChartPreset preset = ExcelDashboardChartPreset.Comparison,
            int? styleId = null, int? colorStyleId = null) {
            ArgumentNullException.ThrowIfNull(chart);
            if (!Enum.IsDefined(typeof(ExcelDashboardChartPreset), preset))
                throw new ArgumentOutOfRangeException(nameof(preset));
            ApplyDashboardAppearance(chart, preset, styleId, colorStyleId);
            return chart;
        }

        /// <summary>Builds a dashboard table and optional chart from a data table.</summary>
        /// <param name="sheet">Worksheet that receives the dashboard.</param>
        /// <param name="data">Tabular data to write.</param>
        /// <param name="options">Dashboard layout and styling options.</param>
        /// <returns>Dashboard table range and generated chart metadata.</returns>
        public static ExcelDashboardResult AddDashboard(this ExcelSheet sheet, DataTable data, ExcelDashboardOptions? options = null) {
            ArgumentNullException.ThrowIfNull(sheet);
            return sheet.AddDashboardCore(data, options);
        }

        /// <summary>Creates a grouped subtotal summary block from a contiguous worksheet data range.</summary>
        /// <param name="sheet">Worksheet that holds the data.</param>
        /// <param name="options">Subtotal generation options.</param>
        /// <returns>Information about the generated summary block.</returns>
        public static ExcelSubtotalResult AddSubtotalSummary(this ExcelSheet sheet, ExcelSubtotalOptions options) {
            ArgumentNullException.ThrowIfNull(sheet);
            return sheet.AddSubtotalSummaryCore(options);
        }

        /// <summary>
        /// Adds a dashboard-ready chart from a range or table and applies preset styling.
        /// </summary>
        /// <param name="sheet">Worksheet that receives the chart.</param>
        /// <param name="options">Dashboard chart options.</param>
        /// <returns>The created chart.</returns>
        public static ExcelChart AddDashboardChart(this ExcelSheet sheet, ExcelDashboardChartOptions options) {
            ArgumentNullException.ThrowIfNull(sheet);
            ArgumentNullException.ThrowIfNull(options);
            if (!Enum.IsDefined(typeof(ExcelDashboardChartPreset), options.Preset)) throw new ArgumentOutOfRangeException(nameof(options.Preset));
            if (options.Row <= 0 || options.Row > ExcelA1.MaxRows) throw new ArgumentOutOfRangeException(nameof(options.Row), "Chart row must be between 1 and the Excel row limit.");
            if (options.Column <= 0 || options.Column > ExcelA1.MaxColumns) throw new ArgumentOutOfRangeException(nameof(options.Column), "Chart column must be between 1 and the Excel column limit.");
            if (options.Width <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(options.Width), "Chart width must be greater than zero.");
            if (options.Height <= OfficeLength.Zero) throw new ArgumentOutOfRangeException(nameof(options.Height), "Chart height must be greater than zero.");
            if (string.IsNullOrWhiteSpace(options.Range) == string.IsNullOrWhiteSpace(options.TableName)) {
                throw new ArgumentException("Provide either Range or TableName.", nameof(options));
            }

            var preset = ResolveDashboardChartPreset(options.Preset);
            var chartType = options.ChartType ?? preset.ChartType;
            OfficeLength width = options.Width ?? OfficeLength.Pixels(preset.WidthPixels);
            OfficeLength height = options.Height ?? OfficeLength.Pixels(preset.HeightPixels);

            ExcelChart chart = !string.IsNullOrWhiteSpace(options.TableName)
                ? sheet.AddChart(chartType, options.TableName!, new OfficeIMO.Excel.ExcelChartOptions { Row = options.Row, Column = options.Column, Width = width, Height = height, Title = options.Title, IncludeCachedData = options.IncludeCachedData })
                : sheet.AddChart(chartType, options.Range!, new OfficeIMO.Excel.ExcelChartOptions { Row = options.Row, Column = options.Column, Width = width, Height = height, HasHeaders = options.HasHeaders, Title = options.Title, IncludeCachedData = options.IncludeCachedData });

            ApplyDashboardAppearance(chart, options.Preset, options.StyleId, options.ColorStyleId);
            return chart;
        }

        private static void ApplyDashboardAppearance(ExcelChart chart, ExcelDashboardChartPreset kind,
            int? styleId, int? colorStyleId) {
            var preset = ResolveDashboardChartPreset(kind);
            chart.ApplyStylePreset(styleId ?? preset.StyleId, colorStyleId ?? preset.ColorStyleId);
            if (preset.HideLegend) {
                chart.HideLegend();
            } else {
                chart.SetLegend(preset.LegendPosition.ToOfficeEnum());
            }

            if (preset.ShowDataLabels) {
                chart.SetDataLabels(
                    showValue: preset.ShowDataLabelValues,
                    showCategoryName: preset.ShowDataLabelCategories,
                    showSeriesName: false,
                    showLegendKey: false,
                    showPercent: preset.ShowDataLabelPercent,
                    position: preset.DataLabelPosition.ToOfficeEnum(),
                    numberFormat: null);
            }

            if (!string.IsNullOrWhiteSpace(chart.Title)) {
                chart.SetTitleTextStyle(fontSize: OfficeLength.Points(12), bold: true);
            }
        }

        private static ResolvedDashboardChartPreset ResolveDashboardChartPreset(ExcelDashboardChartPreset preset) {
            switch (preset) {
                case ExcelDashboardChartPreset.Trend:
                    return new ResolvedDashboardChartPreset(
                        OfficeChartType.Line,
                        widthPixels: 720,
                        heightPixels: 320,
                        styleId: 252,
                        colorStyleId: 11,
                        C.LegendPositionValues.Bottom,
                        hideLegend: false,
                        showDataLabels: false,
                        showDataLabelValues: false,
                        showDataLabelCategories: false,
                        showDataLabelPercent: false,
                        dataLabelPosition: null);
                case ExcelDashboardChartPreset.Contribution:
                    return new ResolvedDashboardChartPreset(
                        OfficeChartType.Pie,
                        widthPixels: 520,
                        heightPixels: 360,
                        styleId: 253,
                        colorStyleId: 12,
                        C.LegendPositionValues.Right,
                        hideLegend: false,
                        showDataLabels: true,
                        showDataLabelValues: false,
                        showDataLabelCategories: true,
                        showDataLabelPercent: true,
                        dataLabelPosition: C.DataLabelPositionValues.BestFit);
                case ExcelDashboardChartPreset.CompactComparison:
                    return new ResolvedDashboardChartPreset(
                        OfficeChartType.BarClustered,
                        widthPixels: 420,
                        heightPixels: 260,
                        styleId: 251,
                        colorStyleId: 10,
                        C.LegendPositionValues.Bottom,
                        hideLegend: true,
                        showDataLabels: true,
                        showDataLabelValues: true,
                        showDataLabelCategories: false,
                        showDataLabelPercent: false,
                        dataLabelPosition: C.DataLabelPositionValues.OutsideEnd);
                default:
                    return new ResolvedDashboardChartPreset(
                        OfficeChartType.ColumnClustered,
                        widthPixels: 640,
                        heightPixels: 360,
                        styleId: 251,
                        colorStyleId: 10,
                        C.LegendPositionValues.Bottom,
                        hideLegend: false,
                        showDataLabels: false,
                        showDataLabelValues: false,
                        showDataLabelCategories: false,
                        showDataLabelPercent: false,
                        dataLabelPosition: null);
            }
        }

        private sealed class ResolvedDashboardChartPreset {
            internal ResolvedDashboardChartPreset(
                OfficeChartType chartType,
                int widthPixels,
                int heightPixels,
                int styleId,
                int colorStyleId,
                C.LegendPositionValues legendPosition,
                bool hideLegend,
                bool showDataLabels,
                bool showDataLabelValues,
                bool showDataLabelCategories,
                bool showDataLabelPercent,
                C.DataLabelPositionValues? dataLabelPosition) {
                ChartType = chartType;
                WidthPixels = widthPixels;
                HeightPixels = heightPixels;
                StyleId = styleId;
                ColorStyleId = colorStyleId;
                LegendPosition = legendPosition;
                HideLegend = hideLegend;
                ShowDataLabels = showDataLabels;
                ShowDataLabelValues = showDataLabelValues;
                ShowDataLabelCategories = showDataLabelCategories;
                ShowDataLabelPercent = showDataLabelPercent;
                DataLabelPosition = dataLabelPosition;
            }

            internal OfficeChartType ChartType { get; }
            internal int WidthPixels { get; }
            internal int HeightPixels { get; }
            internal int StyleId { get; }
            internal int ColorStyleId { get; }
            internal C.LegendPositionValues LegendPosition { get; }
            internal bool HideLegend { get; }
            internal bool ShowDataLabels { get; }
            internal bool ShowDataLabelValues { get; }
            internal bool ShowDataLabelCategories { get; }
            internal bool ShowDataLabelPercent { get; }
            internal C.DataLabelPositionValues? DataLabelPosition { get; }
        }
    }
}
