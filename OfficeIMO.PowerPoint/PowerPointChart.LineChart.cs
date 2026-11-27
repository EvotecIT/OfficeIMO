using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointChart {
        /// <summary>
        /// Sets the grouping mode for a line chart.
        /// </summary>
        /// <param name="grouping">Line chart grouping value.</param>
        /// <returns>The current chart for fluent configuration.</returns>
        public PowerPointChart SetLineChartGrouping(PowerPointChartGrouping grouping) {
            OfficeXmlNode lineChart = FirstDescendant(ChartSpaceNode, "lineChart");
            if (lineChart.IsNone) {
                throw new InvalidOperationException("The chart does not contain a line chart.");
            }
            OfficeXmlNode chartGrouping = lineChart.GetOrPrependChild(CNs, "grouping");
            chartGrouping.SetAttribute("val", grouping.ToOpenXmlToken());
            MarkChartEdited();
            return this;
        }

        /// <summary>The first chart element with this name below <paramref name="element"/> in document order, or none.</summary>
        private static OfficeXmlNode FirstDescendant(OfficeXmlNode element, string localName) {
            foreach (OfficeXmlNode child in element.Children()) {
                if (child.Is(CNs, localName)) return child;
                OfficeXmlNode found = FirstDescendant(child, localName);
                if (found.Exists) return found;
            }
            return default;
        }
    }
}
