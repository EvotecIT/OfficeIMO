using System;
using OfficeIMO.Drawing;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        /// Adds a native chart from the shared OfficeIMO chart contract. Supports all shared chart kinds,
        /// per-series combo kinds, and primary or secondary value axes where the selected families are compatible.
        /// </summary>
        public PowerPointChart AddChart(OfficeChartKind chartKind, OfficeChartData data,
            long left = 0L, long top = 0L, long width = 5486400L, long height = 3200400L,
            PowerPointChartAccessibilityOptions? accessibility = null) {
            if (data == null) throw new ArgumentNullException(nameof(data));
            PowerPointUtils.ValidateSharedChartData(data, chartKind);

            PowerPointChart chart;
            if (chartKind == OfficeChartKind.Scatter) {
                PowerPointScatterChartData scatterData = PowerPointUtils.ToPowerPointScatterChartData(data);
                byte[] workbookBytes = PowerPointUtils.BuildChartWorkbook(scatterData);
                chart = AddChartInternal(workbookBytes, (chartPart, embeddedRelId) => {
                    PowerPointUtils.PopulateChart(chartPart, embeddedRelId, scatterData, PowerPointChartKind.Scatter);
                    PowerPointUtils.ApplySharedChartSeriesStyle(chartPart, data, chartKind);
                }, left, top, width, height);
            } else {
                PowerPointChartData chartData = PowerPointUtils.ToPowerPointChartData(data);
                byte[] workbookBytes = PowerPointUtils.BuildChartWorkbook(chartData);
                chart = AddChartInternal(workbookBytes, (chartPart, embeddedRelId) =>
                    PowerPointUtils.PopulateSharedChart(chartPart, embeddedRelId, data, chartKind),
                    left, top, width, height);
            }

            ApplyChartAccessibility(chart, data, chartKind, accessibility);
            return chart;
        }

        /// <summary>Adds a native shared-contract chart using centimeter measurements.</summary>
        public PowerPointChart AddChartCm(OfficeChartKind chartKind, OfficeChartData data,
            double leftCm, double topCm, double widthCm, double heightCm,
            PowerPointChartAccessibilityOptions? accessibility = null) =>
            AddChart(chartKind, data,
                PowerPointUnits.FromCentimeters(leftCm), PowerPointUnits.FromCentimeters(topCm),
                PowerPointUnits.FromCentimeters(widthCm), PowerPointUnits.FromCentimeters(heightCm), accessibility);

        /// <summary>Adds a native shared-contract chart using inch measurements.</summary>
        public PowerPointChart AddChartInches(OfficeChartKind chartKind, OfficeChartData data,
            double leftInches, double topInches, double widthInches, double heightInches,
            PowerPointChartAccessibilityOptions? accessibility = null) =>
            AddChart(chartKind, data,
                PowerPointUnits.FromInches(leftInches), PowerPointUnits.FromInches(topInches),
                PowerPointUnits.FromInches(widthInches), PowerPointUnits.FromInches(heightInches), accessibility);

        /// <summary>Adds a native shared-contract chart using point measurements.</summary>
        public PowerPointChart AddChartPoints(OfficeChartKind chartKind, OfficeChartData data,
            double leftPoints, double topPoints, double widthPoints, double heightPoints,
            PowerPointChartAccessibilityOptions? accessibility = null) =>
            AddChart(chartKind, data,
                PowerPointUnits.FromPoints(leftPoints), PowerPointUnits.FromPoints(topPoints),
                PowerPointUnits.FromPoints(widthPoints), PowerPointUnits.FromPoints(heightPoints), accessibility);

        private static void ApplyChartAccessibility(PowerPointChart chart, OfficeChartData data,
            OfficeChartKind chartKind, PowerPointChartAccessibilityOptions? options) {
            if (options == null) return;
            if (!string.IsNullOrWhiteSpace(options.Name)) chart.Name = options.Name;
            string? summary = options.DataSummary;
            if (string.IsNullOrWhiteSpace(summary) && options.IncludeDataSummaryInAlternativeText) {
                summary = PowerPointChart.CreateDataSummary(chartKind, data);
            }
            string? alternativeText = options.AlternativeText;
            if (options.IncludeDataSummaryInAlternativeText && !string.IsNullOrWhiteSpace(summary)) {
                alternativeText = string.IsNullOrWhiteSpace(alternativeText)
                    ? "Data summary:" + Environment.NewLine + summary!.Trim()
                    : alternativeText!.Trim() + Environment.NewLine + Environment.NewLine + "Data summary:" +
                      Environment.NewLine + summary!.Trim();
            }
            if (!string.IsNullOrWhiteSpace(alternativeText)) chart.AltText = alternativeText;
        }
    }
}
