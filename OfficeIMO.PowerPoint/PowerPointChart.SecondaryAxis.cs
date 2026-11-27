using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.PowerPoint;

public partial class PowerPointChart {
    /// <summary>Sets independent bounds, tick spacing and appearance, and number format on the existing secondary value axis.</summary>
    public PowerPointChart SetSecondaryValueAxis(OfficeChartValueAxisLayout layout) {
        OfficeOpenXmlChartSecondaryAxis.Apply(ChartSpaceNode.Child(OfficeXmlNamespaces.Chart, "chart"), layout);
        MarkChartEdited();
        return this;
    }
}
