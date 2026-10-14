namespace OfficeIMO.OpenDocument;

public abstract partial class OdfShape {
    /// <summary>Effective start marker. Setting null explicitly disables it, including an inherited marker.</summary>
    public string? StrokeStartMarkerName { get => ReadMarkerName("start"); set => WriteMarkerName("start", value); }
    /// <summary>Effective end marker. Setting null explicitly disables it, including an inherited marker.</summary>
    public string? StrokeEndMarkerName { get => ReadMarkerName("end"); set => WriteMarkerName("end", value); }
    /// <summary>Inherited absolute start-marker width; zero suppresses painting. Null removes the local override.</summary>
    public OdfLength? StrokeStartMarkerWidth { get => ReadMarkerWidth("start"); set => WriteMarkerWidth("start", value); }
    /// <summary>Inherited absolute end-marker width; zero suppresses painting. Null removes the local override.</summary>
    public OdfLength? StrokeEndMarkerWidth { get => ReadMarkerWidth("end"); set => WriteMarkerWidth("end", value); }
    /// <summary>Whether the start marker is centered on its endpoint. Null removes the local override.</summary>
    public bool? StrokeStartMarkerCentered { get => ReadMarkerCenter("start"); set => WriteMarkerCenter("start", value); }
    /// <summary>Whether the end marker is centered on its endpoint. Null removes the local override.</summary>
    public bool? StrokeEndMarkerCentered { get => ReadMarkerCenter("end"); set => WriteMarkerCenter("end", value); }

    private string? ReadMarkerName(string end) {
        string? value = ReadGraphicProperty(OdfNamespaces.Draw + "marker-" + end);
        return string.IsNullOrEmpty(value) ? null : value;
    }
    private void WriteMarkerName(string end, string? name) {
        if (name != null) {
            OdfStyleRepository.ValidateStyleName(name);
            var marker = Document.Styles.FindMarker(name) ?? throw new ArgumentException("Unknown marker '" + name + "'.", nameof(name));
            _ = marker.Geometry;
        }
        EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "marker-" + end, name ?? string.Empty);
    }
    private OdfLength? ReadMarkerWidth(string end) {
        string? value = ReadGraphicProperty(OdfNamespaces.Draw + "marker-" + end + "-width");
        return value == null ? null : OdfLength.Parse(value);
    }
    private void WriteMarkerWidth(string end, OdfLength? value) {
        if (value.HasValue) ValidateMarkerWidth(value.Value);
        EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Draw + "marker-" + end + "-width", value?.ToString());
    }
    private bool? ReadMarkerCenter(string end) {
        string? value = ReadGraphicProperty(OdfNamespaces.Draw + "marker-" + end + "-center");
        return value == null ? null : XmlConvert.ToBoolean(value);
    }
    private void WriteMarkerCenter(string end, bool? value) => EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties",
        OdfNamespaces.Draw + "marker-" + end + "-center", value.HasValue ? XmlConvert.ToString(value.Value) : null);

    internal static double ValidateMarkerWidth(OdfLength value) {
        string text = value.ToString();
        if (text.Length < 3) throw new ArgumentException("Marker width requires an absolute ODF length.", nameof(value));
        string unit = text.Substring(text.Length - 2), number = text.Substring(0, text.Length - 2);
        bool digit = false, dot = false;
        foreach (char c in number) {
            if (c >= '0' && c <= '9') digit = true;
            else if (c == '.' && !dot) dot = true;
            else throw new ArgumentException("Invalid ODF marker width.", nameof(value));
        }
        if (!digit || unit is not ("pt" or "in" or "cm" or "mm" or "pc")) throw new ArgumentException("Invalid ODF marker width.", nameof(value));
        double points = value.ToPoints();
        if (double.IsNaN(points) || double.IsInfinity(points) || points < 0) throw new ArgumentOutOfRangeException(nameof(value));
        return points;
    }
}
