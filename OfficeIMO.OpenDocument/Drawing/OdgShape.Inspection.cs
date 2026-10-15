namespace OfficeIMO.OpenDocument;

public sealed partial class OdgShape {
    /// <summary>Authored image reference, including external targets which inspection never fetches.</summary>
    public string? ImagePath => (string?)Element.Element(OdfNamespaces.Draw + "image")?.Attribute(OdfNamespaces.XLink + "href");

    /// <summary>Detached hyperlink descriptors from this shape's native text story, without accessing targets.</summary>
    public IReadOnlyList<OfficeDocumentModelLink> InspectHyperlinks() => OdfTextTraversal.Paragraphs(TextRoot)
        .SelectMany(OdfTextTraversal.Inlines).Where(element => element.Name == OdfNamespaces.Text + "a")
        .Select((element, index) => new OfficeDocumentModelLink {
            Id = (XmlId ?? Name) + "-link-" + index.ToString(CultureInfo.InvariantCulture),
            Kind = "uri", Uri = (string?)element.Attribute(OdfNamespaces.XLink + "href"), Text = OdfTextCodec.Read(element)
        }).ToArray();
}
