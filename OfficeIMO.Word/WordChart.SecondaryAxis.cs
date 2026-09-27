using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Word;

public partial class WordChart {
    /// <summary>Sets independent bounds, tick spacing, and number format on the existing secondary value axis.</summary>
    public WordChart SetSecondaryValueAxis(OfficeChartValueAxisLayout layout) {
        OfficeOpenXmlChartSecondaryAxis.Apply(_chartPart?.ChartSpace?.GetFirstChild<C.Chart>() ?? _chart, layout);
        _chartPart?.ChartSpace?.Save();
        return this;
    }
}
