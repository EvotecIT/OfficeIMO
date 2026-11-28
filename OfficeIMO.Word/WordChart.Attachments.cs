using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

public partial class WordChart {
    internal List<string> SeriesCategories {
        get => Categories;
        set => Categories = value;
    }

    /// <summary>The chart part's <c>c:chart</c>, or this instance's chart when the part has none.</summary>
    private OfficeXmlNode CurrentChart {
        get {
            OfficeXmlNode space = ChartSpaceNode;
            OfficeXmlNode chart = space.IsNone ? default : space.Child(NsC, "chart");
            return chart.IsNone ? _chart : chart;
        }
    }

    internal void RequireChartAttachment(OfficeXmlNode element) {
        OfficeXmlNode current = CurrentChart;
        if (!current.IsNone) {
            for (OfficeXmlNode ancestor = element.Parent; !ancestor.IsNone; ancestor = ancestor.Parent) {
                if (ancestor.Is(NsC, "chart") && ancestor == current) return;
            }
        }
        throw new InvalidOperationException("The chart item is no longer attached to this chart.");
    }

    internal OfficeSeries ReadAttachedSeries(OfficeXmlNode element) {
        RequireChartAttachment(element);
        OfficeXmlNode group = ContentParent(element);
        if (group.IsNone || !OfficeOpenXmlChartSeriesReader.TryReadKind(group, out var kind))
            throw new NotSupportedException("This native series cannot be inspected.");
        OfficeXmlNode plot = ContentParent(group);
        if (!plot.Is(NsC, "plotArea")) throw new NotSupportedException("The native series has no plot area.");
        var axisGroup = OfficeOpenXmlChartAxisGroups.Read(plot, group);
        OfficeXmlNode scheme = ThemeColorScheme;
        var result = group.Is(NsC, "scatterChart")
            ? OfficeOpenXmlChartSeriesReader.ReadScatter(new[] { element }, scheme, (int)MaxCachedChartPoints)
            : OfficeOpenXmlChartSeriesReader.ReadCategories(new[] { element }, kind, scheme,
                axisGroup, (int)MaxCachedChartPoints);
        return result?.Series.FirstOrDefault()?.Data
            ?? throw new InvalidOperationException("The series has no readable cached values.");
    }
}
