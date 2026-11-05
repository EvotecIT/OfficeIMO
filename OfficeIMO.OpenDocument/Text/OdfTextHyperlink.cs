namespace OfficeIMO.OpenDocument;

/// <summary>A native hyperlink whose target is preserved without network access.</summary>
public sealed class OdfTextHyperlink : OdfTextContent {
    /// <summary>Appends native text and returns this inline.</summary>
    public new OdfTextHyperlink AddText(string text) { base.AddText(text); return this; }
    /// <summary>Transforms stored text and returns the same native receiver.</summary>
    public new OdfTextHyperlink TransformTextCase(OfficeIMO.Drawing.OfficeTextCase textCase, System.Globalization.CultureInfo? culture = null) { base.TransformTextCase(textCase, culture); return this; }

    internal OdfTextHyperlink(OdfDocument document, XElement element, XElement graphic) : base(document, element, graphic, OdfStyleFamily.Text) { }
    /// <summary>Hyperlink target. Relative and absolute references remain as supplied.</summary>
    public string Href {
        get => (string?)Element.Attribute(OdfNamespaces.XLink + "href") ?? string.Empty;
        set { if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Hyperlink target cannot be empty.", nameof(value)); Element.SetAttributeValue(OdfNamespaces.XLink + "href", value); Dirty(); }
    }
    /// <summary>Optional target frame name.</summary>
    public string? TargetFrameName { get => (string?)Element.Attribute(OdfNamespaces.Office + "target-frame-name"); set { Element.SetAttributeValue(OdfNamespaces.Office + "target-frame-name", value); Dirty(); } }
}
