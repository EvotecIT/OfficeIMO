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
    public bool TryGetOfficeSnapshot(out OfficeChartSnapshot snapshot) {
        snapshot = null!;
        try {
            var chart = _chartPart?.ChartSpace?.GetFirstChild<C.Chart>() ?? _chart;
            if (_chartPart == null || chart == null) return false;
            // A formula-linked title without a usable cache cannot be rendered
            // faithfully from the chart part alone.
            if (chart.Descendants<C.Title>().Any(title => title.GetFirstChild<C.ChartText>()?
                .GetFirstChild<C.StringReference>() is C.StringReference reference &&
                reference.Formula != null && string.IsNullOrWhiteSpace(reference.StringCache?.InnerText))) return false;
            // The shared Drawing palette has no Word color-slot remapping metadata.
            // Preserve the native chart and report unsupported projection rather than
            // resolving theme slots through the unmapped document theme.
            if (HasNonIdentityColorSchemeMapping(_document.MainDocumentPartRoot.DocumentSettingsPart?.Settings?
                .GetFirstChild<W.ColorSchemeMapping>())) return false;
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
                OfficeOpenXmlChartSeriesReader.ReadStyle(chart, kind, scheme, textStyle), OfficeOpenXmlChartSeriesReader.ReadLayout(chart, kind, officeData, axisTitleFont, scheme),
                bubbleScale, bubbleMode, OfficeOpenXmlChartRadialLayout.Read(chart));
            if (HasUnclippedExplicitScale(snapshot) ||
                OfficeChartDrawingRenderer.HasUnsupportedAxisUnitBudget(snapshot)) {
                snapshot = null!;
                return false;
            }
            return true;
        } catch {
            snapshot = null!;
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

    private static bool HasUnclippedExplicitScale(OfficeChartSnapshot snapshot) {
        OfficeChartLayout layout = snapshot.Layout;
        foreach (OfficeChartSeries series in snapshot.Data.Series) {
            OfficeChartKind kind = series.RenderKind ?? snapshot.ChartKind;
            bool lineOrArea = kind is OfficeChartKind.Line or OfficeChartKind.LineStacked or OfficeChartKind.LineStacked100 or
                OfficeChartKind.Area or OfficeChartKind.AreaStacked or OfficeChartKind.AreaStacked100;
            bool numericPoints = kind is OfficeChartKind.Scatter or OfficeChartKind.Bubble;
            if (!lineOrArea && !numericPoints) continue;
            // The shared renderer has no plot clipping for these marks yet. A stacked series
            // can cross a bound through its cumulative value even when each source value fits.
            if (kind is OfficeChartKind.LineStacked or OfficeChartKind.LineStacked100 or
                OfficeChartKind.AreaStacked or OfficeChartKind.AreaStacked100 &&
                (layout.VerticalAxisMinimum.HasValue || layout.VerticalAxisMaximum.HasValue)) return true;
            // An unstacked area also paints the polygon down to zero. Without clipping,
            // an explicit range that excludes zero lets its baseline escape the plot.
            if (kind == OfficeChartKind.Area && IsOutside(0d, layout.VerticalAxisMinimum, layout.VerticalAxisMaximum)) return true;
            if (series.Values.Any(value => IsOutside(value, layout.VerticalAxisMinimum, layout.VerticalAxisMaximum))) return true;
            if (numericPoints && (layout.HorizontalAxisMinimum.HasValue || layout.HorizontalAxisMaximum.HasValue) &&
                (series.XValues == null || series.XValues.Any(value => IsOutside(value, layout.HorizontalAxisMinimum, layout.HorizontalAxisMaximum)))) return true;
        }
        return false;
    }

    private static bool IsOutside(double value, double? minimum, double? maximum) =>
        minimum.HasValue && value < minimum.Value || maximum.HasValue && value > maximum.Value;
}
