using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointChart {
    /// <summary>Native pie rotation and doughnut hole geometry, with native defaults when unspecified.</summary>
    public OfficeChartRadialLayout RadialLayout => OfficeOpenXmlChartRadialLayout.Read(GetChartPart().ChartSpace?.GetFirstChild<C.Chart>());

    /// <summary>Sets geometry for an existing native two-dimensional pie or doughnut chart.</summary>
    public PowerPointChart SetRadialLayout(OfficeChartRadialLayout layout) {
        var part = GetChartPart();
        OfficeOpenXmlChartRadialLayout.Apply(part.ChartSpace?.GetFirstChild<C.Chart>(), layout);
        part.ChartSpace!.Save();
        return this;
    }
}
