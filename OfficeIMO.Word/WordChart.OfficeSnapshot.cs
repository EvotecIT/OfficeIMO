using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Word;

public partial class WordChart {
    /// <summary>Reads supported native two-dimensional charts into the shared Drawing contract.</summary>
    /// <remarks>Includes bubble charts and supported category combinations with secondary value axes.
    /// Unqualified families, appearances and cached data return false without modifying the document.</remarks>
    /// <param name="snapshot">The cached chart data and supported presentation metadata.</param>
    /// <returns>True when a complete supported chart projection can be produced.</returns>
    public bool TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot) {
        snapshot = null!;
        try {
            var chart = _chartPart?.ChartSpace?.GetFirstChild<C.Chart>() ?? _chart;
            if (_chartPart == null || chart == null) return false;
            var scheme = _document.MainDocumentPartRoot.ThemePart?.Theme?.ThemeElements?.ColorScheme;
            var data = OfficeOpenXmlChartSeriesReader.ReadPlot(_chartPart, chart, scheme, (int)MaxCachedChartPoints,
                out var kind, out var bubbleScale, out var bubbleMode);
            if (data == null) return false;
            snapshot = new OfficeChartSnapshot(ReadDrawingName(), ReadTitle(chart), kind, data.ToData(), GetWidthPoints(), GetHeightPoints(),
                style: null, OfficeOpenXmlChartSeriesReader.ReadLayout(chart, kind), bubbleScale, bubbleMode, OfficeOpenXmlChartRadialLayout.Read(chart));
            return true;
        } catch {
            snapshot = null!;
            return false;
        }
    }
}
