using OfficeIMO.Drawing;
using OfficeIMO.OpenXml.Internal;
using C = DocumentFormat.OpenXml.Drawing.Charts;
using W = DocumentFormat.OpenXml.Wordprocessing;

namespace OfficeIMO.Word;

public partial class WordChart {
    /// <summary>Reads supported native charts into the shared two-dimensional Drawing contract.</summary>
    /// <remarks>Includes bubble charts and supported category combinations with secondary value axes.
    /// Legacy three-dimensional bar, line, area and pie charts retain their flat cached-data projection.
    /// Unqualified families, appearances and cached data return false without modifying the document.</remarks>
    /// <param name="snapshot">The cached chart data and supported presentation metadata.</param>
    /// <returns>True when a complete supported chart projection can be produced.</returns>
    public bool TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot) =>
        TryGetOfficeSnapshot(null, null, out snapshot, out _);

    /// <summary>Reads the chart at an optional logical extent (point units by default) and reports why an unsupported chart cannot be projected.</summary>
    internal bool TryGetOfficeSnapshot(double? width, double? height, out OfficeChartSnapshot snapshot,
        out string? failureReason) {
        snapshot = null!;
        failureReason = null;
        try {
            var chart = _chartPart?.ChartSpace?.GetFirstChild<C.Chart>() ?? _chart;
            if (_chartPart == null || chart == null) return Fail("The chart part is missing or has no chart.", out failureReason);
            // A formula-linked title without a usable cache cannot be rendered
            // faithfully from the chart part alone.
            if (OfficeOpenXmlChartTextReader.HasUncachedFormulaTitle(chart)) return Fail("The chart title is a formula without a cached value.", out failureReason);
            // The shared Drawing palette has no Word color-slot remapping metadata.
            // Preserve the native chart and report unsupported projection rather than
            // resolving theme slots through the unmapped document theme.
            if (HasNonIdentityColorSchemeMapping(_document.MainDocumentPartRoot.DocumentSettingsPart?.Settings?
                .GetFirstChild<W.ColorSchemeMapping>())) return Fail("The document remaps theme color slots.", out failureReason);
            double renderWidth = width ?? GetWidthPoints();
            double renderHeight = height ?? GetHeightPoints();
            // Preserve the existing flat projection for single legacy 3-D groups. A rejected
            // 2-D projection must never fall back through a less strict legacy reader.
            var groups = chart.PlotArea?.ChildElements.Where(element => element.LocalName.EndsWith("Chart", StringComparison.Ordinal)).Take(10001).ToArray();
            if (groups?.Length > 10000) return Fail("The chart has too many chart groups.", out failureReason);
            if (groups?.Length == 1 && groups[0] is C.Bar3DChart or C.Line3DChart or C.Area3DChart or C.Pie3DChart) {
                if (!TryGetSnapshot(out var legacy) || !OfficeOpenXmlChartSeriesReader.TryReadKind(groups[0], out var legacyKind)) return Fail("The three-dimensional chart has no flat projection.", out failureReason);
                snapshot = new OfficeChartSnapshot(legacy.Name, legacy.Title, legacyKind,
                    new OfficeChartData(legacy.Data.Categories, legacy.Data.Series.Select(series => series.ToOfficeSeries()).ToArray()),
                    global::OfficeIMO.OfficeLength.Points(width ?? legacy.WidthPoints), global::OfficeIMO.OfficeLength.Points(height ?? legacy.HeightPoints), style: null, layout: null, radialLayout: legacy.RadialLayout);
                return true;
            }
            var scheme = _document.MainDocumentPartRoot.ThemePart?.Theme?.ThemeElements?.ColorScheme;
            var data = OfficeOpenXmlChartSeriesReader.ReadPlot(_chartPart, chart, scheme, (int)MaxCachedChartPoints,
                out var kind, out var bubbleScale, out var bubbleMode);
            if (data == null) return Fail("The chart family, series appearance or cached data has no shared chart projection.", out failureReason);
            var textReader = new OfficeOpenXmlChartTextReader(_document.MainDocumentPartRoot.ThemePart?.Theme?.ThemeElements?.FontScheme);
            if (!textReader.TryReadSharedTextStyle(chart, out var textStyle) ||
                !textReader.TryReadAxisTitleTypeface(chart, OfficeOpenXmlChartTextReader.ReadChartDefaultTypeface(chart), out var axisTitleFont)) return Fail("The chart text formatting cannot be projected.", out failureReason);
            OfficeChartData officeData = data.ToData();
            snapshot = new OfficeChartSnapshot(ReadDrawingName(), ReadTitle(chart), kind, officeData, global::OfficeIMO.OfficeLength.Points(renderWidth), global::OfficeIMO.OfficeLength.Points(renderHeight),
                OfficeOpenXmlChartSeriesReader.ReadStyle(_chartPart, chart, kind, scheme, textStyle, data.Categories.Count),
                OfficeOpenXmlChartSeriesReader.ReadLayout(chart, kind, officeData, axisTitleFont, scheme),
                bubbleScale, bubbleMode, OfficeOpenXmlChartRadialLayout.Read(chart));
            if (OfficeChartDrawingRenderer.HasUnsupportedAxisUnitBudget(snapshot)) {
                snapshot = null!;
                return Fail("The axis units produce more ticks than the shared chart renderer draws.", out failureReason);
            }
            return true;
        } catch (Exception exception) {
            snapshot = null!;
            failureReason = exception is NotSupportedException or InvalidDataException
                ? exception.Message
                : exception.GetType().Name + ": " + exception.Message;
            return false;
        }

        static bool Fail(string reason, out string? failure) {
            failure = reason;
            return false;
        }
    }

    private static bool HasNonIdentityColorSchemeMapping(W.ColorSchemeMapping? mapping) => mapping != null && (
        mapping.Background1?.Value != W.ColorSchemeIndexValues.Light1 ||
        mapping.Text1?.Value != W.ColorSchemeIndexValues.Dark1 ||
        mapping.Background2?.Value != W.ColorSchemeIndexValues.Light2 ||
        mapping.Text2?.Value != W.ColorSchemeIndexValues.Dark2 ||
        mapping.Accent1?.Value != W.ColorSchemeIndexValues.Accent1 ||
        mapping.Accent2?.Value != W.ColorSchemeIndexValues.Accent2 ||
        mapping.Accent3?.Value != W.ColorSchemeIndexValues.Accent3 ||
        mapping.Accent4?.Value != W.ColorSchemeIndexValues.Accent4 ||
        mapping.Accent5?.Value != W.ColorSchemeIndexValues.Accent5 ||
        mapping.Accent6?.Value != W.ColorSchemeIndexValues.Accent6 ||
        mapping.Hyperlink?.Value != W.ColorSchemeIndexValues.Hyperlink ||
        mapping.FollowedHyperlink?.Value != W.ColorSchemeIndexValues.FollowedHyperlink);

}
