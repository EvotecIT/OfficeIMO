using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Excel {
    public sealed partial class ExcelChart {
        private static void ClearAutomaticSeriesColor(OfficeXmlNode series) => OfficeOpenXmlChartAutomaticColors.ClearAutomaticSeriesColor(series);
    }
}
