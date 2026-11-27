using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using static OfficeIMO.Excel.OpenXml.ExcelChartXml;

namespace OfficeIMO.Excel;

public sealed partial class ExcelChart {
    /// <summary>Native pie rotation and doughnut hole geometry, with native defaults when unspecified.</summary>
    public OfficeChartRadialLayout RadialLayout => OfficeOpenXmlChartRadialLayout.Read(ChartSpaceNode().Child(C, "chart"));

    /// <summary>Sets geometry for an existing native two-dimensional pie or doughnut chart.</summary>
    public ExcelChart SetRadialLayout(OfficeChartRadialLayout layout) {
        OfficeOpenXmlChartRadialLayout.Apply(ChartSpaceNode().Child(C, "chart"), layout);
        Save();
        return this;
    }
}
