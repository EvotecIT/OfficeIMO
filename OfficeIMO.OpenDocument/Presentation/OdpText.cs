namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed ODP text box frame.</summary>
public sealed class OdpTextBox : OdpShape {
    internal OdpTextBox(OdpPresentation presentation, XElement element) : base(presentation, element) { }
    private XElement TextBox => Element.Element(OdfNamespaces.Draw + "text-box") ?? throw new InvalidDataException("ODP text frame has no draw:text-box.");
    /// <summary>Paragraphs in reading order, including list item paragraphs.</summary>
    public IReadOnlyList<OdpParagraph> Paragraphs => TextBox.Descendants()
        .Where(element => element.Name == OdfNamespaces.Text + "p" || element.Name == OdfNamespaces.Text + "h")
        .Select(element => new OdpParagraph(Presentation, element)).ToList();
    /// <summary>Lists directly contained in the text box.</summary>
    public IReadOnlyList<OdpList> Lists => TextBox.Elements(OdfNamespaces.Text + "list").Select(element => new OdpList(Presentation, element)).ToList();
    /// <summary>Adds a paragraph.</summary>
    public OdpParagraph AddParagraph(string? text = null) {
        var paragraph = new XElement(OdfNamespaces.Text + "p"); OdfTextCodec.Append(paragraph, text); TextBox.Add(paragraph); Dirty(); return new OdpParagraph(Presentation, paragraph);
    }
    /// <summary>Adds an ordered or unordered list.</summary>
    public OdpList AddList(bool ordered = false) {
        string styleName = OdfListStyleStore.Create(Presentation, ordered);
        var list = new XElement(OdfNamespaces.Text + "list", new XAttribute(OdfNamespaces.Text + "style-name", styleName)); TextBox.Add(list); Dirty(); return new OdpList(Presentation, list);
    }
    internal static OdpTextBox Create(OdpPresentation presentation, OdfRect bounds, string? text, string name) {
        var frame = new XElement(OdfNamespaces.Draw + "frame", new XAttribute(OdfNamespaces.Draw + "name", name), new XElement(OdfNamespaces.Draw + "text-box"));
        ApplyBounds(frame, bounds); var result = new OdpTextBox(presentation, frame); if (text != null) result.AddParagraph(text); return result;
    }
}

/// <summary>An XML-backed presentation paragraph.</summary>
public sealed class OdpParagraph {
    private readonly OdpPresentation _presentation;
    private readonly XElement _element;
    internal OdpParagraph(OdpPresentation presentation, XElement element) { _presentation = presentation; _element = element; }
    /// <summary>Decoded paragraph text.</summary>
    public string Text { get => OdfTextCodec.Read(_element); set { OdfTextCodec.Replace(_element, value); Dirty(); } }
    /// <summary>Referenced paragraph style name.</summary>
    public string? StyleName { get => (string?)_element.Attribute(OdfNamespaces.Text + "style-name"); set { _element.SetAttributeValue(OdfNamespaces.Text + "style-name", value); Dirty(); } }
    /// <summary>Inline text runs.</summary>
    public IReadOnlyList<OdpRun> Runs => _element.Descendants(OdfNamespaces.Text + "span").Select(element => new OdpRun(_presentation, element)).ToList();
    /// <summary>Explicit or inherited bold state.</summary>
    public bool? Bold { get => Resolve(style => style.Bold); set => EnsureStyle().Bold = value; }
    /// <summary>Explicit or inherited font size.</summary>
    public OdfLength? FontSize { get => Resolve(style => style.FontSize); set => EnsureStyle().FontSize = value; }
    /// <summary>Adds an inline text run.</summary>
    public OdpRun AddRun(string? text = null) { var span = new XElement(OdfNamespaces.Text + "span"); OdfTextCodec.Append(span, text); _element.Add(span); Dirty(); return new OdpRun(_presentation, span); }
    private OdfStyle EnsureStyle() => _presentation.Styles.EnsureAutomaticStyle(_element, OdfNamespaces.Text + "style-name", OdfStyleFamily.Paragraph, "ofPr");
    private T? Resolve<T>(Func<OdfStyle, T?> selector) where T : struct {
        OdfStyle? style = StyleName == null ? null : _presentation.Styles.Find(OdfStyleFamily.Paragraph, StyleName); if (style == null) return null;
        foreach (OdfStyle candidate in _presentation.Styles.Resolve(style)) { T? value = selector(candidate); if (value.HasValue) return value; } return null;
    }
    private void Dirty() => _presentation.MarkPartDirty("content.xml");
}

/// <summary>An XML-backed presentation inline text run.</summary>
public sealed class OdpRun {
    private readonly OdpPresentation _presentation; private readonly XElement _element;
    internal OdpRun(OdpPresentation presentation, XElement element) { _presentation = presentation; _element = element; }
    /// <summary>Decoded run text.</summary>
    public string Text { get => OdfTextCodec.Read(_element); set { OdfTextCodec.Replace(_element, value); Dirty(); } }
    /// <summary>Referenced text style name.</summary>
    public string? StyleName { get => (string?)_element.Attribute(OdfNamespaces.Text + "style-name"); set { _element.SetAttributeValue(OdfNamespaces.Text + "style-name", value); Dirty(); } }
    /// <summary>Explicit or inherited bold state.</summary>
    public bool? Bold { get => Resolve(style => style.Bold); set => EnsureStyle().Bold = value; }
    /// <summary>Explicit or inherited italic state.</summary>
    public bool? Italic { get => Resolve(style => style.Italic); set => EnsureStyle().Italic = value; }
    /// <summary>Explicit or inherited text color.</summary>
    public OdfColor? Color { get => Resolve(style => style.Color); set => EnsureStyle().Color = value; }
    private OdfStyle EnsureStyle() => _presentation.Styles.EnsureAutomaticStyle(_element, OdfNamespaces.Text + "style-name", OdfStyleFamily.Text, "ofRun");
    private T? Resolve<T>(Func<OdfStyle, T?> selector) where T : struct {
        OdfStyle? style = StyleName == null ? null : _presentation.Styles.Find(OdfStyleFamily.Text, StyleName); if (style == null) return null;
        foreach (OdfStyle candidate in _presentation.Styles.Resolve(style)) { T? value = selector(candidate); if (value.HasValue) return value; } return null;
    }
    private void Dirty() => _presentation.MarkPartDirty("content.xml");
}

/// <summary>An XML-backed presentation list.</summary>
public sealed class OdpList {
    private readonly OdpPresentation _presentation; private readonly XElement _element;
    internal OdpList(OdpPresentation presentation, XElement element) { _presentation = presentation; _element = element; }
    /// <summary>List item paragraphs.</summary>
    public IReadOnlyList<OdpParagraph> Items => _element.Elements(OdfNamespaces.Text + "list-item")
        .SelectMany(item => item.Elements(OdfNamespaces.Text + "p")).Select(element => new OdpParagraph(_presentation, element)).ToList();
    /// <summary>Adds a one-paragraph list item.</summary>
    public OdpParagraph AddItem(string? text = null) {
        var paragraph = new XElement(OdfNamespaces.Text + "p"); OdfTextCodec.Append(paragraph, text);
        _element.Add(new XElement(OdfNamespaces.Text + "list-item", paragraph)); _presentation.MarkPartDirty("content.xml"); return new OdpParagraph(_presentation, paragraph);
    }
}
