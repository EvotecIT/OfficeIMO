using System;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointSlide {
        /// <summary>
        /// Adds a native chart from the shared OfficeIMO chart contract. Supports all shared chart kinds,
        /// per-series combo kinds, and primary or secondary value axes where the selected families are compatible.
        /// </summary>
        public PowerPointChart AddChart(OfficeChartType chartKind,OfficeChartData data,
            OfficeLength? left = null, OfficeLength? top = null, OfficeLength? width = null, OfficeLength? height = null,
            PowerPointChartAccessibilityOptions? accessibility = null) {
            long nativeLeft = (left ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeTop = (top ?? OfficeLength.Emus(0L)).ToEmus();
            long nativeWidth = (width ?? OfficeLength.Emus(5486400L)).ToEmus();
            long nativeHeight = (height ?? OfficeLength.Emus(3200400L)).ToEmus();
            if (data == null) throw new ArgumentNullException(nameof(data));
            byte[] workbookBytes = OfficeOpenXmlChartWriter.BuildWorkbook(data, chartKind);
            PowerPointChart chart = AddChartInternal(workbookBytes, (chartPart, embeddedRelId) =>
                OfficeOpenXmlChartWriter.PopulateSharedChart(chartPart, embeddedRelId, data, chartKind),
                nativeLeft, nativeTop, nativeWidth, nativeHeight);
            ApplyChartAccessibility(chart, data, chartKind, accessibility);
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
