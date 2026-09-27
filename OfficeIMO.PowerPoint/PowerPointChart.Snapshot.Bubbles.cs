using System.Collections.Generic;
using DocumentFormat.OpenXml.Drawing;
using OfficeIMO.OpenXml.Internal;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointChart {
        private static PowerPointChartData? ReadBubbleSeriesData(
            IEnumerable<C.BubbleChartSeries> seriesElements,
            ColorScheme? colorScheme = null, bool forDataUpdate = false) =>
            ProjectSharedSeries(OfficeOpenXmlChartSeriesReader.ReadBubbles(seriesElements,
                colorScheme, PowerPointUtils.MaximumSharedChartPoints, forDataUpdate),
                PowerPointChartSnapshotKind.Bubble);
    }
}