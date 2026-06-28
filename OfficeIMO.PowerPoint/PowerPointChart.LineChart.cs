using System;
using System.Linq;
using DocumentFormat.OpenXml.Packaging;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointChart {
        /// <summary>
        /// Sets the grouping mode for a line chart.
        /// </summary>
        /// <param name="grouping">Open XML line chart grouping value.</param>
        /// <returns>The current chart for fluent configuration.</returns>
        public PowerPointChart SetLineChartGrouping(C.GroupingValues grouping) {
            ChartPart chartPart = GetChartPart();
            C.LineChart lineChart = chartPart.ChartSpace?.Descendants<C.LineChart>().FirstOrDefault()
                ?? throw new InvalidOperationException("The chart does not contain a line chart.");
            C.Grouping chartGrouping = lineChart.GetFirstChild<C.Grouping>() ?? lineChart.PrependChild(new C.Grouping());
            chartGrouping.Val = grouping;
            Save();
            return this;
        }
    }
}
