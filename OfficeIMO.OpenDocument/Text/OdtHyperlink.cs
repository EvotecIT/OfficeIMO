namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed ODT hyperlink. Targets are preserved and never fetched.</summary>
public sealed class OdtHyperlink {
    private readonly OdtDocument _document;
    private readonly XElement _element;
    private readonly string _partPath;

    internal OdtHyperlink(OdtDocument document, XElement element, string partPath = "content.xml") {
        _document = document;
        _element = element;
        _partPath = partPath;
    }

    /// <summary>Decoded display text.</summary>
    public string Text {
        get => OdtTextCodec.Read(_element);
        set { OdtTextCodec.Replace(_element, value); Dirty(); }
    }
    /// <summary>Link target.</summary>
    public string Href {
        get => (string?)_element.Attribute(OdfNamespaces.XLink + "href") ?? string.Empty;
        set {
            if (string.IsNullOrWhiteSpace(value)) throw new ArgumentException("Hyperlink target cannot be empty.", nameof(value));
            _element.SetAttributeValue(OdfNamespaces.XLink + "href", value);
            Dirty();
        }
    }

    private void Dirty() => _document.MarkPartDirty(_partPath);
}
