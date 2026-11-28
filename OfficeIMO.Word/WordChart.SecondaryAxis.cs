using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

public partial class WordChart {
    /// <summary>Sets independent bounds, tick spacing and appearance, and number format on the existing secondary value axis.</summary>
    public WordChart SetSecondaryValueAxis(OfficeChartValueAxisLayout layout) {
        OfficeOpenXmlChartSecondaryAxis.Apply(CurrentChart, layout);
        SaveChartSpace();
        return this;
    }
}
