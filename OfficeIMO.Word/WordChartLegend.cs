using C = DocumentFormat.OpenXml.Drawing.Charts;

namespace OfficeIMO.Word;

/// <summary>An editable legend attached to a Word chart.</summary>
public sealed class WordChartLegend {
    private readonly C.Legend _element;
    internal WordChartLegend(WordChart chart, C.Legend element) { Chart = chart; _element = element; }
    /// <summary>The owning chart.</summary>
    public WordChart Chart { get; }
    /// <summary>The legend position.</summary>
    public OfficeChartLegendPosition Position {
        get { RequireAttached(); return (_element.LegendPosition?.Val?.Value ?? C.LegendPositionValues.Right).ToOfficeEnum(); }
        set { RequireAttached(); _element.LegendPosition = new C.LegendPosition { Val = value.ToOpenXml() }; }
    }
    /// <summary>Whether the legend overlays the plot.</summary>
    public bool Overlay {
        get { RequireAttached(); return _element.GetFirstChild<C.Overlay>()?.Val?.Value ?? false; }
        set { RequireAttached(); _element.GetFirstChild<C.Overlay>()?.Remove(); _element.AddChild(new C.Overlay { Val = value }, true); }
    }
    /// <summary>Removes this legend from its chart.</summary>
    public void Remove() { RequireAttached(); _element.Remove(); }
    private void RequireAttached() => Chart.RequireChartAttachment(_element);
}
