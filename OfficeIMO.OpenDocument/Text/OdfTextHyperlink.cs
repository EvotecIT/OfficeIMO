namespace OfficeIMO.OpenDocument;

/// <summary>A native hyperlink whose target is preserved without network access.</summary>
public sealed class OdfTextHyperlink : OdfTextContent {
    internal OdfTextHyperlink(OdfDocument document, XElement element, XElement graphic) : base(document, element, graphic, OdfStyleFamily.Text) { }
    /// <summary>Hyperlink target. Relative and absolute references remain as supplied.</summary>
    public string Href {
        get => (string?)Element.Attribute(OdfNamespaces.XLink + "href") ?? string.Empty;
        set { if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Hyperlink target cannot be empty.", nameof(value)); Element.SetAttributeValue(OdfNamespaces.XLink + "href", value); Dirty(); }
    }
    /// <summary>Optional target frame name.</summary>
    public string? TargetFrameName { get => (string?)Element.Attribute(OdfNamespaces.Office + "target-frame-name"); set { Element.SetAttributeValue(OdfNamespaces.Office + "target-frame-name", value); Dirty(); } }
}
