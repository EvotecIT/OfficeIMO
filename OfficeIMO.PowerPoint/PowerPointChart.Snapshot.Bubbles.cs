using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint {
    public partial class PowerPointChart {
        private static PowerPointChartData? ReadBubbleSeriesData(
            IEnumerable<OfficeXmlNode> seriesElements,
            OfficeXmlNode colorScheme = default, bool forDataUpdate = false,
            OfficeOpenXmlChartSeriesReader.SourceVisibility? visibility = null) =>
            ProjectSharedSeries(OfficeOpenXmlChartSeriesReader.ReadBubbles(seriesElements,
                colorScheme, PowerPointUtils.MaximumSharedChartPoints, forDataUpdate, visibility: visibility),
                PowerPointChartSnapshotKind.Bubble);
    }
}
