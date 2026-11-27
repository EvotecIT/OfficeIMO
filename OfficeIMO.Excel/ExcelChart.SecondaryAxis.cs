using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel;

public partial class ExcelChart {
    /// <summary>Sets independent bounds, tick spacing and appearance, and number format on the existing secondary value axis.</summary>
    public ExcelChart SetSecondaryValueAxis(OfficeChartValueAxisLayout layout) {
        var part = ChartPartRef();
        OfficeOpenXmlChartSecondaryAxis.Apply(part.Root.Child(C, "chart"), layout);
        Save();
        return this;
    }
}
