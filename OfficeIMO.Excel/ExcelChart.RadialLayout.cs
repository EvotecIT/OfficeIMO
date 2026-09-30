using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Excel;

public sealed partial class ExcelChart {
    /// <summary>Native pie rotation and doughnut hole geometry, with native defaults when unspecified.</summary>
    public OfficeChartRadialLayout RadialLayout => OfficeOpenXmlChartRadialLayout.Read(GetChartPart().ChartSpace?.GetFirstChild<C.Chart>());

    /// <summary>Sets geometry for an existing native two-dimensional pie or doughnut chart.</summary>
    public ExcelChart SetRadialLayout(OfficeChartRadialLayout layout) {
        OfficeOpenXmlChartRadialLayout.Apply(GetChartPart().ChartSpace?.GetFirstChild<C.Chart>(), layout);
        Save();
        return this;
    }
}
