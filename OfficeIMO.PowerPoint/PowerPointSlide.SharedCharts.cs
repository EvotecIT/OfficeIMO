using System;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>Adds an editable native chart with the format's default title, size and placement.</summary>
        public PowerPointChart AddChart(OfficeChartType chartKind, OfficeChartData data) =>
            AddChart(chartKind, data, null);

        /// <summary>
        /// Adds a native chart from the shared OfficeIMO chart contract. Supports all shared chart kinds,
        /// per-series combo kinds, and primary or secondary value axes where the selected families are compatible.
        /// </summary>
        public PowerPointChart AddChart(OfficeChartType chartKind, OfficeChartData data, OfficeChartOptions? options) {
            var placement = options as PowerPointChartOptions;
            long nativeLeft = (placement?.Left ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeTop = (placement?.Top ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeWidth = (options?.Width ?? OfficeLength.Emus(5486400L)).ToEmus();
            long nativeHeight = (options?.Height ?? OfficeLength.Emus(3200400L)).ToEmus();
            if (nativeWidth <= 0) throw new ArgumentOutOfRangeException(nameof(options), "Chart width must be positive.");
            if (nativeHeight <= 0) throw new ArgumentOutOfRangeException(nameof(options), "Chart height must be positive.");
            if (data == null) throw new ArgumentNullException(nameof(data));
            byte[] workbookBytes = OfficeOpenXmlChartWriter.BuildWorkbook(data, chartKind);
            PowerPointChart chart = AddChartInternal(workbookBytes, (chartPart, embeddedRelId) =>
                OfficeOpenXmlChartWriter.PopulateSharedChart(chartPart, embeddedRelId, data, chartKind),
                nativeLeft, nativeTop, nativeWidth, nativeHeight);
            if (options?.Title != null) chart.SetTitle(options.Title);
            ApplyChartAccessibility(chart, data, chartKind, placement?.Accessibility);
            return chart;
        }

        private static void ApplyChartAccessibility(PowerPointChart chart, OfficeChartData data,
            OfficeChartType chartKind, PowerPointChartAccessibilityOptions? options) {
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
