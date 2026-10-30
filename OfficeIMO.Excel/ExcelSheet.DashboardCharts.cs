using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Excel {
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

    public partial class ExcelSheet {
        /// <summary>
        /// Adds a dashboard-ready chart from a range or table and applies preset styling.
        /// </summary>
        /// <param name="options">Dashboard chart options.</param>
        /// <returns>The created chart.</returns>
        public ExcelChart AddDashboardChart(ExcelDashboardChartOptions options) {
            if (options == null) throw new ArgumentNullException(nameof(options));
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
            int styleId = options.StyleId ?? preset.StyleId;
            int colorStyleId = options.ColorStyleId ?? preset.ColorStyleId;

            ExcelChart chart = !string.IsNullOrWhiteSpace(options.TableName)
                ? AddChart(chartType, options.TableName!, new OfficeIMO.Excel.ExcelChartOptions { Row = options.Row, Column = options.Column, Width = width, Height = height, Title = options.Title, IncludeCachedData = options.IncludeCachedData })
                : AddChart(chartType, options.Range!, new OfficeIMO.Excel.ExcelChartOptions { Row = options.Row, Column = options.Column, Width = width, Height = height, HasHeaders = options.HasHeaders, Title = options.Title, IncludeCachedData = options.IncludeCachedData });

            chart.ApplyStylePreset(styleId, colorStyleId);
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

            if (!string.IsNullOrWhiteSpace(options.Title)) {
                chart.SetTitleTextStyle(fontSize: OfficeLength.Points(12), bold: true);
            }

            return chart;
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
