using DocumentFormat.OpenXml;
using OfficeIMO.OpenXml.Internal;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Word;

public partial class WordChart {
    internal List<string> SeriesCategories {
        get => Categories;
        set => Categories = value;
    }

    internal void RequireChartAttachment(OpenXmlElement element) {
        var current = _chartPart?.ChartSpace?.GetFirstChild<C.Chart>() ?? _chart;
        if (current == null || !element.Ancestors<C.Chart>().Any(chart => ReferenceEquals(chart, current)))
            throw new InvalidOperationException("The chart item is no longer attached to this chart.");
    }

    internal OfficeSeries ReadAttachedSeries(OpenXmlCompositeElement element) {
        RequireChartAttachment(element);
        if (element.Parent is not OpenXmlCompositeElement group ||
            !OfficeOpenXmlChartSeriesReader.TryReadKind(group, out var kind))
            throw new NotSupportedException("This native series cannot be inspected.");
        var scheme = _document.MainDocumentPartRoot.ThemePart?.Theme?.ThemeElements?.ColorScheme;
        var result = element is C.ScatterChartSeries scatter
            ? OfficeOpenXmlChartSeriesReader.ReadScatter(new[] { scatter }, scheme, (int)MaxCachedChartPoints)
            : OfficeOpenXmlChartSeriesReader.ReadCategories(new[] { element }, kind, scheme,
                OfficeIMO.Drawing.OfficeChartAxisGroup.Primary, (int)MaxCachedChartPoints);
        return result?.Series.FirstOrDefault()?.Data
            ?? throw new InvalidOperationException("The series has no readable cached values.");
    }
}
