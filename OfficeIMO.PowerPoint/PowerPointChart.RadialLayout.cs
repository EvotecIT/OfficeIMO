using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointChart {
    /// <summary>Native pie rotation and doughnut hole geometry, with native defaults when unspecified.</summary>
    public OfficeChartRadialLayout RadialLayout =>
        OfficeOpenXmlChartRadialLayout.Read(ChartSpaceNode.Child(OfficeXmlNamespaces.Chart, "chart"));

    /// <summary>Sets geometry for an existing native two-dimensional pie or doughnut chart.</summary>
    public PowerPointChart SetRadialLayout(OfficeChartRadialLayout layout) {
        OfficeOpenXmlChartRadialLayout.Apply(ChartSpaceNode.Child(OfficeXmlNamespaces.Chart, "chart"), layout);
        MarkChartEdited();
        return this;
    }
}
