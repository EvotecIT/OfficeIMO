namespace OfficeIMO.OpenDocument;

/// <summary>An XML-backed ODT paragraph or heading.</summary>
public sealed class OdtParagraph {
    private readonly OdtDocument _document;
    private readonly XElement _element;
    private readonly string _partPath;

    internal OdtParagraph(OdtDocument document, XElement element, string partPath = "content.xml") {
        _document = document;
        _element = element;
        _partPath = partPath;
    }

    /// <summary>Plain text with ODF spaces, tabs, and line breaks decoded.</summary>
    public string Text {
        get => OdfTextCodec.Read(_element);
        set {
            OdfTextCodec.Replace(_element, value);
            Dirty();
        }
    }

    /// <summary>Referenced paragraph style name.</summary>
    public string? StyleName {
        get => (string?)_element.Attribute(OdfNamespaces.Text + "style-name");
        set {
            _element.SetAttributeValue(OdfNamespaces.Text + "style-name", value);
            Dirty();
        }
    }

    /// <summary>True when this block is a heading.</summary>
    public bool IsHeading => _element.Name == OdfNamespaces.Text + "h";

    /// <summary>Heading outline level, or null for a normal paragraph.</summary>
    public int? HeadingLevel {
        get {
            if (!IsHeading) return null;
            return int.TryParse((string?)_element.Attribute(OdfNamespaces.Text + "outline-level"), NumberStyles.Integer,
                CultureInfo.InvariantCulture, out int level) ? level : 1;
        }
        set {
            if (!value.HasValue) throw new ArgumentNullException(nameof(value));
            if (value < 1 || value > 10) throw new ArgumentOutOfRangeException(nameof(value));
            _element.Name = OdfNamespaces.Text + "h";
            _element.SetAttributeValue(OdfNamespaces.Text + "outline-level", value.Value);
            Dirty();
        }
    }

    /// <summary>Inline text spans in this paragraph.</summary>
    public IReadOnlyList<OdtSpan> Spans => _element.Descendants(OdfNamespaces.Text + "span")
        .Select(element => new OdtSpan(_document, element, _partPath)).ToList();

    /// <summary>Hyperlinks in this paragraph.</summary>
    public IReadOnlyList<OdtHyperlink> Hyperlinks => _element.Descendants(OdfNamespaces.Text + "a")
        .Select(element => new OdtHyperlink(_document, element, _partPath)).ToList();

    /// <summary>Controls whether this paragraph starts on a new page.</summary>
    public bool PageBreakBefore {
        get => ResolveStyleValue(style => style.BreakBefore) == "page";
        set {
            OdfStyle style = EnsureStyle();
            style.BreakBefore = value ? "page" : null;
        }
    }

    /// <summary>Explicit or inherited bold state.</summary>
    public bool? Bold {
        get => ResolveStyleValue(style => style.Bold);
        set => EnsureStyle().Bold = value;
    }

    /// <summary>Explicit or inherited italic state.</summary>
    public bool? Italic {
        get => ResolveStyleValue(style => style.Italic);
        set => EnsureStyle().Italic = value;
    }

    /// <summary>Explicit or inherited font size.</summary>
    public OdfLength? FontSize {
        get => ResolveStyleValue(style => style.FontSize);
        set => EnsureStyle().FontSize = value;
    }

    /// <summary>Explicit or inherited text color.</summary>
    public OdfColor? Color {
        get => ResolveStyleValue(style => style.Color);
        set => EnsureStyle().Color = value;
    }

    /// <summary>Appends plain text while encoding ODF whitespace semantics.</summary>
    public OdtParagraph AddText(string text) {
        OdfTextCodec.Append(_element, text);
        Dirty();
        return this;
    }

    /// <summary>Appends a styled text span.</summary>
    public OdtSpan AddSpan(string? text = null) {
        var element = new XElement(OdfNamespaces.Text + "span");
        OdfTextCodec.Append(element, text);
        _element.Add(element);
        Dirty();
        return new OdtSpan(_document, element, _partPath);
    }

    /// <summary>Appends a hyperlink without resolving or fetching its target.</summary>
    public OdtHyperlink AddHyperlink(string text, string href) {
        if (string.IsNullOrWhiteSpace(href)) throw new ArgumentException("Hyperlink target cannot be empty.", nameof(href));
        var element = new XElement(OdfNamespaces.Text + "a",
            new XAttribute(OdfNamespaces.XLink + "type", "simple"),
            new XAttribute(OdfNamespaces.XLink + "href", href));
        OdfTextCodec.Append(element, text);
        _element.Add(element);
        Dirty();
        return new OdtHyperlink(_document, element, _partPath);
    }

    /// <summary>Appends a collapsed bookmark.</summary>
    public OdtParagraph AddBookmark(string name) {
        ValidateBookmarkName(name);
        _element.Add(new XElement(OdfNamespaces.Text + "bookmark", new XAttribute(OdfNamespaces.Text + "name", name)));
        Dirty();
        return this;
    }

    /// <summary>Appends a bookmark range start marker.</summary>
    public OdtParagraph AddBookmarkStart(string name) {
        ValidateBookmarkName(name);
        _element.Add(new XElement(OdfNamespaces.Text + "bookmark-start", new XAttribute(OdfNamespaces.Text + "name", name)));
        Dirty();
        return this;
    }

    /// <summary>Appends a bookmark range end marker.</summary>
    public OdtParagraph AddBookmarkEnd(string name) {
        ValidateBookmarkName(name);
        _element.Add(new XElement(OdfNamespaces.Text + "bookmark-end", new XAttribute(OdfNamespaces.Text + "name", name)));
        Dirty();
        return this;
    }

    /// <summary>Appends an inline or paragraph-anchored image.</summary>
    public OdtImage AddImage(byte[] data, string fileName, OdfLength width, OdfLength height,
        OdtImageAnchor anchor = OdtImageAnchor.Inline) {
        OdtImage image = OdtImage.Create(_document, data, fileName, width, height, anchor);
        _element.Add(image.Element);
        Dirty();
        return image;
    }

    internal XElement Element => _element;

    private OdfStyle EnsureStyle() => _document.Styles.EnsureAutomaticStyle(
        _element, OdfNamespaces.Text + "style-name", OdfStyleFamily.Paragraph, "ofP", _partPath);

    private T? ResolveStyleValue<T>(Func<OdfStyle, T?> selector) where T : struct {
        OdfStyle? style = StyleName == null ? null : _document.Styles.Find(OdfStyleFamily.Paragraph, StyleName);
        if (style == null) return null;
        foreach (OdfStyle candidate in _document.Styles.Resolve(style)) {
            T? value = selector(candidate);
            if (value.HasValue) return value;
        }
        return null;
    }

    private string? ResolveStyleValue(Func<OdfStyle, string?> selector) {
        OdfStyle? style = StyleName == null ? null : _document.Styles.Find(OdfStyleFamily.Paragraph, StyleName);
        if (style == null) return null;
        foreach (OdfStyle candidate in _document.Styles.Resolve(style)) {
            string? value = selector(candidate);
            if (value != null) return value;
        }
        return null;
    }

    private static void ValidateBookmarkName(string name) {
        if (string.IsNullOrWhiteSpace(name)) throw new ArgumentException("Bookmark name cannot be empty.", nameof(name));
    }

    private void Dirty() => _document.MarkPartDirty(_partPath);
}
