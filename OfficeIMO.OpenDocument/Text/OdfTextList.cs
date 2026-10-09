namespace OfficeIMO.OpenDocument;

/// <summary>A native Draw text list with preserved levels, numbering, and continuation declarations.</summary>
public sealed class OdfTextList {
    private readonly OdfDocument _document;
    private readonly XElement _element;
    private readonly XElement _graphic;
    internal OdfTextList(OdfDocument document, XElement element, XElement graphic) { _document = document; _element = element; _graphic = graphic; }
    /// <summary>Referenced list style name. A nested list without a style inherits its enclosing list's style.</summary>
    public string? StyleName { get => (string?)_element.Attribute(OdfNamespaces.Text + "style-name"); set { _element.SetAttributeValue(OdfNamespaces.Text + "style-name", value); _document.MarkPartDirty(_document.GetPartPath(_graphic)); } }
    /// <summary>Whether the effective native list-level style is numbered.</summary>
    public bool IsOrdered => OdfListStyleStore.IsOrdered(_document,
        _element.AncestorsAndSelf().Where(e => e.Name == OdfNamespaces.Text + "list").Select(e => (string?)e.Attribute(OdfNamespaces.Text + "style-name")).FirstOrDefault(n => n != null),
        partPath: _document.GetPartPath(_graphic), level: _element.AncestorsAndSelf().Count(e => e.Name == OdfNamespaces.Text + "list"));
    /// <summary>Direct list items; headers remain preserved separately in native XML.</summary>
    public IReadOnlyList<OdfTextListItem> Items => _element.Elements(OdfNamespaces.Text + "list-item").Select(e => new OdfTextListItem(_document, e, _graphic)).ToList();
    /// <summary>Adds an item containing one paragraph.</summary>
    public OdfTextListItem AddItem(string? text = null) {
        var paragraph = new XElement(OdfNamespaces.Text + "p"); OdfTextCodec.Append(paragraph, text);
        var item = new XElement(OdfNamespaces.Text + "list-item", paragraph); _element.Add(item); _document.MarkPartDirty(_document.GetPartPath(_graphic)); return new OdfTextListItem(_document, item, _graphic);
    }
    /// <summary>Returns a detached copy of the native list.</summary>
    public XElement ToXml() => new XElement(_element);
}

/// <summary>A native list item containing paragraphs and optionally nested lists.</summary>
public sealed class OdfTextListItem {
    private readonly OdfDocument _document;
    private readonly XElement _element;
    private readonly XElement _graphic;
    internal OdfTextListItem(OdfDocument document, XElement element, XElement graphic) { _document = document; _element = element; _graphic = graphic; }
    /// <summary>Optional nonnegative numbering restart for this item. Null uses the list counter.</summary>
    public long? StartValue {
        get => _element.Attribute(OdfNamespaces.Text + "start-value") is XAttribute attribute ? long.Parse(attribute.Value, CultureInfo.InvariantCulture) : null;
        set {
            if (value < 0) throw new ArgumentOutOfRangeException(nameof(value));
            _element.SetAttributeValue(OdfNamespaces.Text + "start-value", value?.ToString(CultureInfo.InvariantCulture));
            _document.MarkPartDirty(_document.GetPartPath(_graphic));
        }
    }
    /// <summary>Direct paragraphs and headings in this item.</summary>
    public IReadOnlyList<OdfTextParagraph> Paragraphs => _element.Elements().Where(OdfTextTraversal.IsParagraph).Select(e => new OdfTextParagraph(_document, e, _graphic)).ToList();
    /// <summary>Direct nested lists.</summary>
    public IReadOnlyList<OdfTextList> Lists => _element.Elements(OdfNamespaces.Text + "list").Select(e => new OdfTextList(_document, e, _graphic)).ToList();
    /// <summary>Adds a paragraph after existing content.</summary>
    public OdfTextParagraph AddParagraph(string? text = null) {
        var paragraph = new XElement(OdfNamespaces.Text + "p"); OdfTextCodec.Append(paragraph, text); _element.Add(paragraph); _document.MarkPartDirty(_document.GetPartPath(_graphic)); return new OdfTextParagraph(_document, paragraph, _graphic);
    }
    /// <summary>Adds a nested list with an independent native style at its actual nesting level, up to level ten.</summary>
    public OdfTextList AddList(bool ordered = false) {
        int level = _element.Ancestors().Count(e => e.Name == OdfNamespaces.Text + "list") + 1;
        string style = OdfListStyleStore.Create(_document, ordered, _document.GetPartPath(_graphic), listLevel: level);
        var list = new XElement(OdfNamespaces.Text + "list", new XAttribute(OdfNamespaces.Text + "style-name", style));
        _element.Add(list); _document.MarkPartDirty(_document.GetPartPath(_graphic)); return new OdfTextList(_document, list, _graphic);
    }
    /// <summary>Returns a detached copy of the native item, including numbering overrides.</summary>
    public XElement ToXml() => new XElement(_element);
}
