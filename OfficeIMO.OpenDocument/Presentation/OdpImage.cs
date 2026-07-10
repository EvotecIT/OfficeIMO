namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed ODP image frame.</summary>
public sealed class OdpImage : OdpShape {
    internal OdpImage(OdpPresentation presentation, XElement element) : base(presentation, element) { }
    /// <summary>Package-relative image path.</summary>
    public string Path => (string?)Element.Element(OdfNamespaces.Draw + "image")?.Attribute(OdfNamespaces.XLink + "href") ?? string.Empty;
    /// <summary>Image crop insets stored as <c>fo:clip</c>.</summary>
    public OdfInsets? Crop {
        get {
            string? lexical = (string?)GetGraphicProperties()?.Attribute(OdfNamespaces.Fo + "clip");
            if (lexical == null || !lexical.StartsWith("rect(", StringComparison.Ordinal) || !lexical.EndsWith(")", StringComparison.Ordinal)) return null;
            string[] values = lexical.Substring(5, lexical.Length - 6).Split(new[] { ' ', ',' }, StringSplitOptions.RemoveEmptyEntries);
            return values.Length == 4 ? new OdfInsets(OdfLength.Parse(values[0]), OdfLength.Parse(values[1]), OdfLength.Parse(values[2]), OdfLength.Parse(values[3])) : (OdfInsets?)null;
        }
        set {
            string? lexical = value.HasValue ? "rect(" + value.Value.Top + " " + value.Value.Right + " " + value.Value.Bottom + " " + value.Value.Left + ")" : null;
            EnsureGraphicStyle().SetProperty(OdfNamespaces.Style + "graphic-properties", OdfNamespaces.Fo + "clip", lexical);
        }
    }
    /// <summary>Returns a defensive copy of the embedded image bytes.</summary>
    public byte[] GetImageBytes() => Presentation.GetPackageEntryBytes(Path);
    internal static OdpImage Create(OdpPresentation presentation, byte[] data, string fileName, OdfRect bounds, string name) {
        string path = OdfImageStore.Add(presentation, data, fileName);
        var frame = new XElement(OdfNamespaces.Draw + "frame", new XAttribute(OdfNamespaces.Draw + "name", name),
            new XElement(OdfNamespaces.Draw + "image", new XAttribute(OdfNamespaces.XLink + "href", path),
                new XAttribute(OdfNamespaces.XLink + "type", "simple"), new XAttribute(OdfNamespaces.XLink + "show", "embed"),
                new XAttribute(OdfNamespaces.XLink + "actuate", "onLoad")));
        ApplyBounds(frame, bounds); return new OdpImage(presentation, frame);
    }
    private XElement? GetGraphicProperties() {
        string? styleName = (string?)Element.Attribute(OdfNamespaces.Draw + "style-name");
        return styleName == null ? null : Presentation.Styles.Find(OdfStyleFamily.Graphic, styleName)?.Element.Element(OdfNamespaces.Style + "graphic-properties");
    }
}
