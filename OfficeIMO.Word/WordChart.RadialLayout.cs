using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

public partial class WordChart {
    /// <summary>Native pie rotation and doughnut hole geometry, with native defaults when unspecified.</summary>
    public OfficeChartRadialLayout RadialLayout => OfficeOpenXmlChartRadialLayout.Read(_chartPart?.ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>() ?? _chart);

    /// <summary>Sets geometry for an existing native two-dimensional pie or doughnut chart.</summary>
    public WordChart SetRadialLayout(OfficeChartRadialLayout layout) {
        OfficeOpenXmlChartRadialLayout.Apply(_chartPart?.ChartSpace?.GetFirstChild<DocumentFormat.OpenXml.Drawing.Charts.Chart>() ?? _chart, layout);
        _chartPart?.ChartSpace?.Save();
        return this;
    }
}
