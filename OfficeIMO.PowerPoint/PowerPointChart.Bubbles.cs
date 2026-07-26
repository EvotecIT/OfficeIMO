using System;
using System.Linq;
using DocumentFormat.OpenXml;
using DocumentFormat.OpenXml.Packaging;
using OfficeIMO.Drawing;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointChart {
        /// <summary>
        /// Sets the native bubble scale and whether bubble values represent area or width.
        /// </summary>
        /// <param name="scalePercent">Bubble diameter scale from zero through 300 percent.</param>
        /// <param name="sizeMode">Whether bubble values represent area or width.</param>
        public PowerPointChart SetBubbleSizing(
            uint scalePercent, OfficeChartBubbleSizeMode sizeMode) {
            if (scalePercent > 300U) {
                throw new ArgumentOutOfRangeException(nameof(scalePercent),
                    "Bubble scale must be from zero through 300 percent.");
            }
            if (!Enum.IsDefined(typeof(OfficeChartBubbleSizeMode), sizeMode)) {
                throw new ArgumentOutOfRangeException(nameof(sizeMode));
            }

            ChartPart chartPart = GetChartPart();
            C.PlotArea? plotArea = chartPart.ChartSpace?
                .GetFirstChild<C.Chart>()?
                .GetFirstChild<C.PlotArea>();
            C.BubbleChart[] bubbleCharts = plotArea?
                .Elements<C.BubbleChart>()
                .ToArray() ?? Array.Empty<C.BubbleChart>();
            if (bubbleCharts.Length == 0) {
                throw new NotSupportedException(
                    "Bubble sizing can only be set on a bubble chart.");
            }

            foreach (C.BubbleChart bubbleChart in bubbleCharts) {
                ApplyBubbleSizing(bubbleChart, scalePercent, sizeMode);
            }

            Save();
            return this;
        }

        private static void ApplyBubbleSizing(C.BubbleChart bubbleChart,
            uint scalePercent, OfficeChartBubbleSizeMode sizeMode) {
            C.BubbleScale scale = bubbleChart.GetFirstChild<C.BubbleScale>() ??
                new C.BubbleScale();
            scale.Val = scalePercent;
            if (scale.Parent == null) {
                OpenXmlElement? insertBefore =
                    bubbleChart.GetFirstChild<C.ShowNegativeBubbles>() ??
                    (OpenXmlElement?)bubbleChart.GetFirstChild<C.SizeRepresents>() ??
                    bubbleChart.GetFirstChild<C.AxisId>();
                if (insertBefore == null) bubbleChart.Append(scale);
                else bubbleChart.InsertBefore(scale, insertBefore);
            }

            C.SizeRepresents represents =
                bubbleChart.GetFirstChild<C.SizeRepresents>() ??
                new C.SizeRepresents();
            represents.Val = sizeMode == OfficeChartBubbleSizeMode.Width
                ? C.SizeRepresentsValues.Width
                : C.SizeRepresentsValues.Area;
            if (represents.Parent == null) {
                OpenXmlElement? insertBefore =
                    bubbleChart.GetFirstChild<C.AxisId>();
                if (insertBefore == null) bubbleChart.Append(represents);
                else bubbleChart.InsertBefore(represents, insertBefore);
            }
        }
    }
}
