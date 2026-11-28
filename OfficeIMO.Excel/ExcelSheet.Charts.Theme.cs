using OfficeIMO.Excel.OpenXml;

namespace OfficeIMO.Excel {
    public partial class ExcelSheet {
        private static void ApplyDefaultChartTheme(ExcelPartRef chartPart) =>
            OfficeIMO.OpenXml.Internal.OfficeOpenXmlChartThemeWriter.Apply(chartPart.AsChartPart(), OfficeTheme.Default);
    }
}
