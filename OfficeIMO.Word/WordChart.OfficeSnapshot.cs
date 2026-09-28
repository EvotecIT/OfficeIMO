using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Word;

public partial class WordChart {
    /// <summary>Reads supported native charts into the shared two-dimensional Drawing contract.</summary>
    /// <remarks>Includes bubble charts and supported category combinations with secondary value axes.
    /// Legacy three-dimensional bar, line, area and pie charts retain their flat cached-data projection.
    /// Unqualified families, appearances and cached data return false without modifying the document.</remarks>
    /// <param name="snapshot">The cached chart data and supported presentation metadata.</param>
    /// <returns>True when a complete supported chart projection can be produced.</returns>
    public bool TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot) {
        snapshot = null!;
        try {
            var chart = _chartPart?.ChartSpace?.GetFirstChild<C.Chart>() ?? _chart;
            if (_chartPart == null || chart == null) return false;
            // Preserve the existing flat projection for single legacy 3-D groups. A rejected
            // 2-D projection must never fall back through a less strict legacy reader.
            var groups = chart.PlotArea?.ChildElements.Where(element => element.LocalName.EndsWith("Chart", StringComparison.Ordinal)).Take(10001).ToArray();
            if (groups?.Length > 10000) return false;
            if (groups?.Length == 1 && groups[0] is C.Bar3DChart or C.Line3DChart or C.Area3DChart or C.Pie3DChart) {
                if (!TryGetSnapshot(out var legacy) || !OfficeOpenXmlChartSeriesReader.TryReadKind(groups[0], out var legacyKind)) return false;
                snapshot = new OfficeChartSnapshot(legacy.Name, legacy.Title, legacyKind,
                    new OfficeChartData(legacy.Data.Categories, legacy.Data.Series.Select(series => series.ToOfficeSeries()).ToArray()),
                    legacy.WidthPoints, legacy.HeightPoints, style: null, layout: null, radialLayout: legacy.RadialLayout);
                return true;
            }
            var scheme = _document.MainDocumentPartRoot.ThemePart?.Theme?.ThemeElements?.ColorScheme;
            var data = OfficeOpenXmlChartSeriesReader.ReadPlot(_chartPart, chart, scheme, (int)MaxCachedChartPoints,
                out var kind, out var bubbleScale, out var bubbleMode);
            if (data == null) return false;
            var textReader = new OfficeOpenXmlChartTextReader(_document.MainDocumentPartRoot.ThemePart?.Theme?.ThemeElements?.FontScheme);
            if (!textReader.TryReadSharedTextStyle(chart, out var textStyle) ||
                !textReader.TryReadAxisTitleTypeface(chart, OfficeOpenXmlChartTextReader.ReadChartDefaultTypeface(chart), out var axisTitleFont)) return false;
            OfficeChartData officeData = data.ToData();
            snapshot = new OfficeChartSnapshot(ReadDrawingName(), ReadTitle(chart), kind, officeData, GetWidthPoints(), GetHeightPoints(),
                OfficeOpenXmlChartSeriesReader.ReadStyle(_chartPart, chart, kind, scheme, textStyle, data.Categories.Count),
                OfficeOpenXmlChartSeriesReader.ReadLayout(chart, kind, officeData, axisTitleFont, scheme),
                bubbleScale, bubbleMode, OfficeOpenXmlChartRadialLayout.Read(chart));
            return true;
        } catch {
            snapshot = null!;
            return false;
        }
    }

}
