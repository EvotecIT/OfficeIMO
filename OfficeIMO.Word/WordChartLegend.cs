using OfficeIMO.Core.Internal.OpenXml;
using OfficeIMO.OpenXml.Internal;

namespace OfficeIMO.Word;

/// <summary>An editable legend attached to a Word chart.</summary>
public sealed class WordChartLegend {
    private const int NsC = OfficeXmlNamespaces.Chart;
    private readonly OfficeXmlNode _element;
    internal WordChartLegend(WordChart chart, OfficeXmlNode element) { Chart = chart; _element = element; }
    /// <summary>The owning chart.</summary>
    public WordChart Chart { get; }
    /// <summary>The legend position.</summary>
    public OfficeChartLegendPosition Position {
        get {
            RequireAttached();
            OfficeXmlNode position = _element.Child(NsC, "legendPos");
            string token = (position.IsNone ? null : position.Val()) ?? "r";
            return WordOpenXmlValueTypeExtensions.TryParseOpenXmlValue(token, out OfficeChartLegendPosition value)
                ? value : throw new FormatException("The legend position '" + token + "' is not a valid value.");
        }
        set {
            RequireAttached();
            OfficeChartXml.SetOrdered(_element, _element.CreateVal(NsC, "legendPos", value.ToOpenXmlToken()), CLegendView.Order);
        }
    }
    /// <summary>Whether the legend overlays the plot.</summary>
    public bool Overlay {
        get {
            RequireAttached();
            OfficeXmlNode overlay = _element.Child(NsC, "overlay");
            return (overlay.IsNone ? null : overlay.BoolVal()) ?? false;
        }
        set {
            RequireAttached();
            OfficeXmlNode overlay = _element.Child(NsC, "overlay");
            if (!overlay.IsNone) overlay.Remove();
            OfficeChartXml.SetOrdered(_element, _element.CreateVal(NsC, "overlay", OfficeChartXml.Text(value)), CLegendView.Order);
        }
    }
    /// <summary>Removes this legend from its chart.</summary>
    public void Remove() { RequireAttached(); _element.Remove(); }
    private void RequireAttached() => Chart.RequireChartAttachment(_element);
}
