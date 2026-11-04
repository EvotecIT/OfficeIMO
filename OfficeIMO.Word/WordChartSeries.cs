using DocumentFormat.OpenXml;
using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Word;

/// <summary>Chart series data, optionally attached to an editable native chart.</summary>
public sealed class WordChartSeries {
    private readonly OfficeSeries? _snapshot;
    private readonly OpenXmlCompositeElement? _element;

    internal WordChartSeries(OfficeSeries snapshot) => _snapshot = snapshot;
    internal WordChartSeries(WordChart chart, OpenXmlCompositeElement element) {
        Chart = chart;
        _element = element;
    }

    /// <summary>The owning chart, or null for detached inspection data.</summary>
    public WordChart? Chart { get; }

    /// <summary>The series display name.</summary>
    public string Name {
        get => _snapshot?.Name ?? RequireAttached().GetFirstChild<C.SeriesText>()?
            .Descendants<C.NumericValue>().FirstOrDefault()?.Text ?? string.Empty;
        set {
            if (value == null) throw new ArgumentNullException(nameof(value));
            var element = RequireAttached();
            var text = element.GetFirstChild<C.SeriesText>()
                ?? throw new InvalidOperationException("The series has no native title.");
            text.RemoveAllChildren();
            text.Append(new C.NumericValue(value));
        }
    }

    /// <summary>Y values or category values.</summary>
    public IReadOnlyList<double> Values => ToOfficeSeries().Values;
    /// <summary>Optional X values for scatter charts.</summary>
    public IReadOnlyList<double>? XValues => ToOfficeSeries().XValues;
    /// <summary>Optional explicit series color.</summary>
    public OfficeColor? Color => ToOfficeSeries().Color;
    /// <summary>Optional explicit point colors.</summary>
    public IReadOnlyList<OfficeColor?>? PointColors => ToOfficeSeries().PointColors;
    /// <summary>Point fill, hatch, and outline overrides.</summary>
    public IReadOnlyList<OfficeIMO.Drawing.OfficeChartPointStyle?>? PointStyles => ToOfficeSeries().PointStyles;

    /// <summary>Reads cached data and appearance through the shared native chart reader.</summary>
    public OfficeSeries ToOfficeSeries() => _snapshot ?? Chart!.ReadAttachedSeries(RequireAttached());

    /// <summary>Removes this native series. Detached inspection data cannot be edited.</summary>
    public void Remove() => RequireAttached().Remove();

    private OpenXmlCompositeElement RequireAttached() {
        if (Chart == null || _element == null)
            throw new InvalidOperationException("Detached chart inspection data cannot be edited.");
        Chart.RequireChartAttachment(_element);
        return _element;
    }
}
