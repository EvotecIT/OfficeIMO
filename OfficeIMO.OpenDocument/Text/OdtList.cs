namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed ordered or unordered ODT list.</summary>
public sealed class OdtList {
    private readonly OdtDocument _document;
    private readonly XElement _element;

    internal OdtList(OdtDocument document, XElement element) {
        _document = document;
        _element = element;
    }

    /// <summary>List items in source order.</summary>
    public IReadOnlyList<OdtListItem> Items => _element.Elements(OdfNamespaces.Text + "list-item")
        .Select(element => new OdtListItem(_document, element)).ToList();

    /// <summary>Adds an item containing one paragraph.</summary>
    public OdtListItem AddItem(string? text = null) {
        var paragraph = new XElement(OdfNamespaces.Text + "p");
        OdfTextCodec.Append(paragraph, text);
        var item = new XElement(OdfNamespaces.Text + "list-item", paragraph);
        _element.Add(item);
        _document.MarkPartDirty("content.xml");
        return new OdtListItem(_document, item);
    }
}

/// <summary>An XML-backed ODT list item.</summary>
public sealed class OdtListItem {
    private readonly OdtDocument _document;
    private readonly XElement _element;

    internal OdtListItem(OdtDocument document, XElement element) {
        _document = document;
        _element = element;
    }

    /// <summary>Paragraphs directly contained by this item.</summary>
    public IReadOnlyList<OdtParagraph> Paragraphs => _element.Elements()
        .Where(element => element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h")
        .Select(element => new OdtParagraph(_document, element)).ToList();

    /// <summary>Adds another paragraph to this item.</summary>
    public OdtParagraph AddParagraph(string? text = null) {
        var paragraph = new XElement(OdfNamespaces.Text + "p");
        OdfTextCodec.Append(paragraph, text);
        _element.Add(paragraph);
        _document.MarkPartDirty("content.xml");
        return new OdtParagraph(_document, paragraph);
    }
}
