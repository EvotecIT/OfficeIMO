namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed named start/end marker definition in common styles.</summary>
public sealed class OdfMarker {
    private readonly OdfDocument _document;
    private readonly XElement _element;
    internal OdfMarker(OdfDocument document, XElement element) { _document = document; _element = element; }
    /// <summary>Name used by shape marker references.</summary>
    public string Name => (string?)_element.Attribute(OdfNamespaces.Draw + "name") ?? string.Empty;
    /// <summary>Reads or replaces marker geometry while retaining unrelated definition metadata.</summary>
    public OdfMarkerGeometry Geometry {
        get => new OdfMarkerGeometry((string?)_element.Attribute(OdfNamespaces.Svg + "viewBox") ?? throw new InvalidDataException("Marker '" + Name + "' has no view box."),
            (string?)_element.Attribute(OdfNamespaces.Svg + "d") ?? throw new InvalidDataException("Marker '" + Name + "' has no path."));
        set {
            if (value == null) throw new ArgumentNullException(nameof(value));
            WriteGeometry(_element, value); _document.MarkPartDirty("styles.xml");
        }
    }
    internal static void WriteGeometry(XElement element, OdfMarkerGeometry geometry) {
        element.SetAttributeValue(OdfNamespaces.Svg + "viewBox", geometry.AuthoringViewBox);
        element.SetAttributeValue(OdfNamespaces.Svg + "d", geometry.PathData);
    }
}
