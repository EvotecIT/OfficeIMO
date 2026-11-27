using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

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

            OfficeXmlNode plotArea = ChartSpaceNode.Child(CNs, "chart").Child(CNs, "plotArea");
            List<OfficeXmlNode> bubbleCharts = plotArea.Children(CNs, "bubbleChart").ToList();
            if (bubbleCharts.Count == 0) {
                throw new NotSupportedException(
                    "Bubble sizing can only be set on a bubble chart.");
            }

            foreach (OfficeXmlNode bubbleChart in bubbleCharts) {
                ApplyBubbleSizing(bubbleChart, scalePercent, sizeMode);
            }

            MarkChartEdited();
            return this;
        }

        private static void ApplyBubbleSizing(OfficeXmlNode bubbleChart,
            uint scalePercent, OfficeChartBubbleSizeMode sizeMode) {
            OfficeXmlNode scale = bubbleChart.Child(CNs, "bubbleScale");
            if (scale.IsNone) {
                scale = InsertBeforeOrAppend(bubbleChart, bubbleChart.CreateElement(CNs, "bubbleScale"),
                    FirstChildOf(bubbleChart, "showNegBubbles", "sizeRepresents", "axId"));
            }
            scale.SetUInt32Attribute("val", scalePercent);

            OfficeXmlNode represents = bubbleChart.Child(CNs, "sizeRepresents");
            if (represents.IsNone) {
                represents = InsertBeforeOrAppend(bubbleChart, bubbleChart.CreateElement(CNs, "sizeRepresents"),
                    bubbleChart.Child(CNs, "axId"));
            }
            represents.SetAttribute("val", sizeMode == OfficeChartBubbleSizeMode.Width ? "w" : "area");
        }
    }
}
