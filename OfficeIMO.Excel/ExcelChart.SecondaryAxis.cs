using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Excel;

public partial class ExcelChart {
    /// <summary>Sets independent bounds, tick spacing and appearance, and number format on the existing secondary value axis.</summary>
    public ExcelChart SetSecondaryValueAxis(OfficeChartValueAxisLayout layout) {
        var part = GetChartPart();
        OfficeOpenXmlChartSecondaryAxis.Apply(part.ChartSpace?.GetFirstChild<C.Chart>(), layout);
        part.ChartSpace!.Save();
        return this;
    }
}
