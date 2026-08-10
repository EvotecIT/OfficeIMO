namespace OfficeIMO.OpenDocument;

/// <summary>Horizontal alignment for an ODT paragraph.</summary>
public enum OdtParagraphAlignment {
    /// <summary>Aligns content to the logical start edge.</summary>
    Start = 0,
    /// <summary>Centers content.</summary>
    Center = 1,
    /// <summary>Aligns content to the logical end edge.</summary>
    End = 2,
    /// <summary>Justifies content on both edges.</summary>
    Justify = 3,
    /// <summary>Aligns content to the physical left edge.</summary>
    Left = 4,
    /// <summary>Aligns content to the physical right edge.</summary>
    Right = 5
}

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

    /// <summary>
    /// Direct inline nodes in document order. Use this syntax view when mixed plain text,
    /// spans, links, images, or bookmark markers must be processed without flattening.
    /// </summary>
    public IReadOnlyList<OdtInlineNode> InlineNodes => OdtInlineNode.Read(_document, _element, _partPath);

    /// <summary>Embedded image frames in this paragraph.</summary>
    public IReadOnlyList<OdtImage> Images => _element.Descendants(OdfNamespaces.Draw + "frame")
        .Where(element => element.Element(OdfNamespaces.Draw + "image") != null)
        .Select(element => new OdtImage(_document, element, _partPath)).ToList();

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

    /// <summary>Explicit or inherited underline state.</summary>
    public bool? Underline {
        get => ResolveStyleValue(style => style.Underline);
        set => EnsureStyle().Underline = value;
    }
    /// <summary>Whether the effective underline uses a non-solid ODF decoration style.</summary>
    public bool UsesNonSolidUnderlineStyle =>
        ResolveStyleValue(style => style.UsesNonSolidUnderlineStyle) == true;

    /// <summary>Explicit or inherited strike-through state.</summary>
    public bool? StrikeThrough {
        get => ResolveStyleValue(style => style.StrikeThrough);
        set => EnsureStyle().StrikeThrough = value;
    }
    /// <summary>Whether the effective line-through uses a non-solid ODF decoration style.</summary>
    public bool UsesNonSolidLineThroughStyle =>
        ResolveStyleValue(style => style.UsesNonSolidLineThroughStyle) == true;

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

    /// <summary>Explicit or inherited text background color.</summary>
    public OdfColor? TextBackgroundColor {
        get {
            OdfStyle? style = StyleName == null ? null : _document.Styles.FindInPart(
                OdfStyleFamily.Paragraph, StyleName, _partPath);
            return _document.Styles.ResolveTextBackgroundColor(style);
        }
        set => EnsureStyle().TextBackgroundColor = value;
    }

    /// <summary>Explicit or inherited font family.</summary>
    public string? FontFamily {
        get => ResolveStyleValue(style => style.FontFamily);
        set => EnsureStyle().FontFamily = value;
    }

    /// <summary>Explicit or inherited paragraph background color.</summary>
    public OdfColor? BackgroundColor {
        get {
            OdfStyle? style = StyleName == null ? null : _document.Styles.FindInPart(
                OdfStyleFamily.Paragraph, StyleName, _partPath);
            return _document.Styles.ResolveBackgroundColor(style);
        }
        set => EnsureStyle().BackgroundColor = value;
    }

    /// <summary>Explicit or inherited horizontal paragraph alignment.</summary>
    public OdtParagraphAlignment? Alignment {
        get => ParseAlignment(ResolveStyleValue(style => style.TextAlign));
        set => EnsureStyle().TextAlign = FormatAlignment(value);
    }

    /// <summary>Effective ODF writing-mode token, such as <c>lr-tb</c> or <c>rl-tb</c>.</summary>
    public string? WritingMode {
        get => ResolveStyleValue(style => style.WritingMode);
        set => EnsureStyle().WritingMode = value;
    }

    /// <summary>Whether the effective horizontal paragraph writing mode is right-to-left.</summary>
    public bool IsRightToLeft => string.Equals(WritingMode, "rl", StringComparison.OrdinalIgnoreCase)
        || string.Equals(WritingMode, "rl-tb", StringComparison.OrdinalIgnoreCase);

    /// <summary>Effective paragraph line height, including absolute and percentage values.</summary>
    public OdfLength? LineHeight {
        get => ResolveStyleValue(style => style.LineHeight);
        set => EnsureStyle().LineHeight = value;
    }

    /// <summary>Explicit or inherited paragraph start indentation.</summary>
    public OdfLength? IndentStart {
        get => ResolveStyleValue(style => style.MarginLeft);
        set => EnsureStyle().MarginLeft = value;
    }

    /// <summary>Explicit or inherited paragraph end indentation.</summary>
    public OdfLength? IndentEnd {
        get => ResolveStyleValue(style => style.MarginRight);
        set => EnsureStyle().MarginRight = value;
    }

    /// <summary>Explicit or inherited first-line indentation.</summary>
    public OdfLength? FirstLineIndent {
        get => ResolveStyleValue(style => style.TextIndent);
        set => EnsureStyle().TextIndent = value;
    }

    /// <summary>Explicit or inherited spacing above the paragraph.</summary>
    public OdfLength? SpaceAbove {
        get => ResolveStyleValue(style => style.MarginTop);
        set => EnsureStyle().MarginTop = value;
    }

    /// <summary>Explicit or inherited spacing below the paragraph.</summary>
    public OdfLength? SpaceBelow {
        get => ResolveStyleValue(style => style.MarginBottom);
        set => EnsureStyle().MarginBottom = value;
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
        OdfStyle? style = StyleName == null ? null : _document.Styles.FindInPart(OdfStyleFamily.Paragraph, StyleName, _partPath);
        if (style == null) return null;
        foreach (OdfStyle candidate in _document.Styles.Resolve(style)) {
            T? value = selector(candidate);
            if (value.HasValue) return value;
        }
        return null;
    }

    private string? ResolveStyleValue(Func<OdfStyle, string?> selector) {
        OdfStyle? style = StyleName == null ? null : _document.Styles.FindInPart(OdfStyleFamily.Paragraph, StyleName, _partPath);
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

    private static OdtParagraphAlignment? ParseAlignment(string? value) {
        switch (value?.ToLowerInvariant()) {
            case "start": return OdtParagraphAlignment.Start;
            case "left": return OdtParagraphAlignment.Left;
            case "center": return OdtParagraphAlignment.Center;
            case "right": return OdtParagraphAlignment.Right;
            case "end": return OdtParagraphAlignment.End;
            case "justify": return OdtParagraphAlignment.Justify;
            default: return null;
        }
    }

    private static string? FormatAlignment(OdtParagraphAlignment? value) {
        switch (value) {
            case OdtParagraphAlignment.Start: return "start";
            case OdtParagraphAlignment.Left: return "left";
            case OdtParagraphAlignment.Center: return "center";
            case OdtParagraphAlignment.Right: return "right";
            case OdtParagraphAlignment.End: return "end";
            case OdtParagraphAlignment.Justify: return "justify";
            default: return null;
        }
    }

    private void Dirty() => _document.MarkPartDirty(_partPath);
}
