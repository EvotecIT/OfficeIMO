using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointChart {
    /// <summary>Sets independent bounds, tick spacing, and number format on the existing secondary value axis.</summary>
    public PowerPointChart SetSecondaryValueAxis(OfficeChartValueAxisLayout layout) {
        var part = GetChartPart();
        OfficeOpenXmlChartSecondaryAxis.Apply(part.ChartSpace?.GetFirstChild<C.Chart>(), layout);
        part.ChartSpace!.Save();
        return this;
    }
}
