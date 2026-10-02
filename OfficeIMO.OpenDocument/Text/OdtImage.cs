namespace OfficeIMO.OpenDocument;

/// <summary>Image anchoring modes supported by the ODT writer.</summary>
public enum OdtImageAnchor {
    /// <summary>Image participates in text flow as a character.</summary>
    Inline,
    /// <summary>Image is anchored to its containing paragraph.</summary>
    Paragraph
}

/// <summary>An XML-backed ODT image frame.</summary>
public sealed class OdtImage {
    private readonly OdtDocument _document;
    private readonly string _partPath;
    private Func<XElement>? _materializeForEdit;

    internal OdtImage(OdtDocument document, XElement element, string partPath = "content.xml", Func<XElement>? materializeForEdit = null) {
        _document = document;
        _partPath = partPath;
        Element = element;
        _materializeForEdit = materializeForEdit;
    }

    /// <summary>Package-relative image path.</summary>
    public string Path => (string?)Element.Element(OdfNamespaces.Draw + "image")?.Attribute(OdfNamespaces.XLink + "href") ?? string.Empty;
    /// <summary>Frame width.</summary>
    public OdfLength Width {
        get => OdfLength.Parse((string?)Element.Attribute(OdfNamespaces.Svg + "width") ?? "0cm");
        set { EnsureMaterialized().SetAttributeValue(OdfNamespaces.Svg + "width", value.ToString()); Dirty(); }
    }
    /// <summary>Frame height.</summary>
    public OdfLength Height {
        get => OdfLength.Parse((string?)Element.Attribute(OdfNamespaces.Svg + "height") ?? "0cm");
        set { EnsureMaterialized().SetAttributeValue(OdfNamespaces.Svg + "height", value.ToString()); Dirty(); }
    }
    /// <summary>Image anchor mode.</summary>
    public OdtImageAnchor Anchor {
        get => (string?)Element.Attribute(OdfNamespaces.Text + "anchor-type") == "paragraph" ? OdtImageAnchor.Paragraph : OdtImageAnchor.Inline;
        set { EnsureMaterialized().SetAttributeValue(OdfNamespaces.Text + "anchor-type", value == OdtImageAnchor.Paragraph ? "paragraph" : "as-char"); Dirty(); }
    }

    /// <summary>Returns a defensive copy of the embedded image bytes.</summary>
    public byte[] GetImageBytes() => _document.GetPackageEntryBytes(Path);

    internal XElement Element { get; private set; }

    internal static OdtImage Create(OdtDocument document, byte[] data, string fileName, OdfLength width, OdfLength height, OdtImageAnchor anchor) {
        string path = OdfImageStore.Add(document, data, fileName);
        int index = document.TextBody.Descendants(OdfNamespaces.Draw + "frame").Count() + 1;
        var frame = new XElement(OdfNamespaces.Draw + "frame",
            new XAttribute(OdfNamespaces.Draw + "name", "Image" + index.ToString(CultureInfo.InvariantCulture)),
            new XAttribute(OdfNamespaces.Text + "anchor-type", anchor == OdtImageAnchor.Paragraph ? "paragraph" : "as-char"),
            new XAttribute(OdfNamespaces.Svg + "width", width.ToString()),
            new XAttribute(OdfNamespaces.Svg + "height", height.ToString()),
            new XElement(OdfNamespaces.Draw + "image",
                new XAttribute(OdfNamespaces.XLink + "href", path),
                new XAttribute(OdfNamespaces.XLink + "type", "simple"),
                new XAttribute(OdfNamespaces.XLink + "show", "embed"),
                new XAttribute(OdfNamespaces.XLink + "actuate", "onLoad")));
        return new OdtImage(document, frame);
    }

    private XElement EnsureMaterialized() {
        if (_materializeForEdit != null) {
            Element = _materializeForEdit();
            _materializeForEdit = null;
        }
        return Element;
    }

    private void Dirty() => _document.MarkPartDirty(_partPath);
}
