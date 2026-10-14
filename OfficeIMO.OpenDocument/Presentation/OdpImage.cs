namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed ODP image frame.</summary>
public sealed class OdpImage : OdpShape {
    internal OdpImage(OdpPresentation presentation, XElement element) : base(presentation, element) { }
    /// <summary>Package-relative image path.</summary>
    public string Path => (string?)Element.Element(OdfNamespaces.Draw + "image")?.Attribute(OdfNamespaces.XLink + "href") ?? string.Empty;
    /// <summary>Image crop insets in intrinsic image lengths, stored as <c>fo:clip</c>.</summary>
    /// <remarks>Negative offsets add empty space. Null removes the local override; explicit zero insets disable inherited cropping.</remarks>
    public OdfInsets? Crop {
        get => ReadImageCrop();
        set => WriteImageCrop(value);
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
}
