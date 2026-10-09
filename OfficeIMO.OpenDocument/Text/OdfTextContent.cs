namespace OfficeIMO.OpenDocument;

/// <summary>Native ODF paragraph or inline content. Edits stay in the owning document XML.</summary>
public abstract class OdfTextContent {
    internal readonly OdfDocument Document;
    internal readonly XElement Element;
    internal readonly XElement Graphic;
    private readonly OdfStyleFamily _family;

    internal OdfTextContent(OdfDocument document, XElement element, XElement graphic, OdfStyleFamily family) {
        Document = document; Element = element; Graphic = graphic; _family = family;
    }

    /// <summary>Decoded text. Assignment replaces this element's child content, including nested formatting and fields.</summary>
    public string Text { get => OdfTextCodec.Read(Element); set { OdfTextCodec.Replace(Element, value); Dirty(); } }
    /// <summary>Referenced native style name.</summary>
    public string? StyleName { get => (string?)Element.Attribute(OdfNamespaces.Text + "style-name"); set { Element.SetAttributeValue(OdfNamespaces.Text + "style-name", value); Dirty(); } }
    /// <summary>Ordered inline syntax. Text values are snapshots; typed wrappers edit the original XML.</summary>
    public IReadOnlyList<OdfTextNode> InlineNodes => OdfTextNode.Read(Document, Element, Graphic);
    /// <summary>Styled runs in this content's text story, including nested runs.</summary>
    public IReadOnlyList<OdfTextRun> Runs => OdfTextTraversal.Inlines(Element).Where(e => e.Name == OdfNamespaces.Text + "span")
        .Select(e => new OdfTextRun(Document, e, Graphic)).ToList();
    /// <summary>Hyperlinks in this content's text story, including links inside runs.</summary>
    public IReadOnlyList<OdfTextHyperlink> Hyperlinks => OdfTextTraversal.Inlines(Element).Where(e => e.Name == OdfNamespaces.Text + "a")
        .Select(e => new OdfTextHyperlink(Document, e, Graphic)).ToList();
    /// <summary>Basic page, date, and time fields in this content's text story.</summary>
    public IReadOnlyList<OdfTextField> Fields => OdfTextTraversal.Inlines(Element).Where(e => OdfTextField.IsField(e.Name))
        .Select(e => new OdfTextField(Document, e)).ToList();

    /// <summary>Effective bold state, including inline, paragraph, and graphic inheritance.</summary>
    public bool? Bold { get => Resolve(s => s.Bold); set => EnsureStyle().Bold = value; }
    /// <summary>Effective italic state.</summary>
    public bool? Italic { get => Resolve(s => s.Italic); set => EnsureStyle().Italic = value; }
    /// <summary>Effective underline state.</summary>
    public bool? Underline { get => Resolve(s => s.Underline); set => EnsureStyle().Underline = value; }
    /// <summary>Effective native underline style.</summary>
    public OdfTextDecorationStyle? UnderlineStyle { get => Resolve(s => s.UnderlineStyle); set => EnsureStyle().UnderlineStyle = value; }
    /// <summary>Effective native underline line count.</summary>
    public OdfTextDecorationType? UnderlineType { get => Resolve(s => s.UnderlineType); set => EnsureStyle().UnderlineType = value; }
    /// <summary>Effective strike-through state.</summary>
    public bool? StrikeThrough { get => Resolve(s => s.StrikeThrough); set => EnsureStyle().StrikeThrough = value; }
    /// <summary>Effective font size.</summary>
    public OdfLength? FontSize { get => Resolve(s => s.FontSize); set => EnsureStyle().FontSize = value; }
    /// <summary>Effective font family.</summary>
    public string? FontFamily { get => Styles.Select(s => s.FontFamily).FirstOrDefault(v => v != null); set => EnsureStyle().FontFamily = value; }
    /// <summary>Effective text color.</summary>
    public OdfColor? Color { get => Resolve(s => s.Color); set => EnsureStyle().Color = value; }
    /// <summary>Effective text opacity from zero to one, including inline, paragraph and graphic inheritance.</summary>
    /// <remarks>Null means no resolved declaration. Assignment edits this element's local text style; null removes its override without changing inherited styles. Rendering follows the qualified foreground and list-label profile.</remarks>
    /// <exception cref="InvalidDataException">The nearest imported opacity declaration is invalid or its aliases conflict.</exception>
    /// <exception cref="ArgumentOutOfRangeException">The assigned value is not finite or is outside zero to one.</exception>
    public double? TextOpacity {
        get => Resolve(s => s.TextOpacity);
        set { OdfOpacity.Validate(value); EnsureStyle().TextOpacity = value; }
    }
    /// <summary>Effective text background. An explicit transparent override stops inheritance.</summary>
    public OdfColor? BackgroundColor {
        get { foreach (OdfStyle style in Styles) if (style.TryGetTextBackgroundColor(out OdfColor? color)) return color; return null; }
        set => EnsureStyle().TextBackgroundColor = value;
    }
    /// <summary>Effective baseline placement.</summary>
    public OdfTextPosition? TextPosition { get => Resolve(s => s.TextPosition); set => EnsureStyle().TextPosition = value; }
    /// <summary>Effective display-time case transformation.</summary>
    public OdfTextTransform? TextTransform { get => Resolve(s => s.TextTransform); set => EnsureStyle().TextTransform = value; }
    /// <summary>Effective small-cap formatting.</summary>
    public bool? SmallCaps { get => Resolve(s => s.SmallCaps); set => EnsureStyle().SmallCaps = value; }

    /// <summary>Appends text using ODF space, tab, and line-break encoding.</summary>
    public void AddText(string text) { OdfTextCodec.Append(Element, text); Dirty(); }
    /// <summary>Appends a styled run, retaining existing nodes.</summary>
    public OdfTextRun AddRun(string? text = null) {
        var span = new XElement(OdfNamespaces.Text + "span"); OdfTextCodec.Append(span, text);
        Element.Add(span); Dirty(); return new OdfTextRun(Document, span, Graphic);
    }
    /// <summary>Appends a hyperlink without fetching or resolving its target.</summary>
    public OdfTextHyperlink AddHyperlink(string text, string href) {
        if (string.IsNullOrWhiteSpace(href)) throw new ArgumentException("Hyperlink target cannot be empty.", nameof(href));
        if (Element.AncestorsAndSelf().Any(e => e.Name == OdfNamespaces.Text + "a")) throw new NotSupportedException("ODF hyperlinks cannot be nested.");
        var link = new XElement(OdfNamespaces.Text + "a", new XAttribute(OdfNamespaces.XLink + "type", "simple"), new XAttribute(OdfNamespaces.XLink + "href", href));
        OdfTextCodec.Append(link, text); Element.Add(link); Dirty(); return new OdfTextHyperlink(Document, link, Graphic);
    }
    /// <summary>Appends a basic native field with cached display text.</summary>
    public OdfTextField AddField(OdfTextFieldKind kind, string? displayText = null) {
        XElement field = OdfTextField.Create(kind, displayText); Element.Add(field); Dirty(); return new OdfTextField(Document, field);
    }
    /// <summary>Transforms stored text while retaining inline XML and excluding annotations and embedded objects.</summary>
    public void TransformTextCase(OfficeIMO.Drawing.OfficeTextCase textCase, CultureInfo? culture = null) {
        OdfTextCodec.TransformTextCase(Element, textCase, culture); Dirty();
    }
    /// <summary>Returns a detached copy of this native element.</summary>
    public XElement ToXml() => new XElement(Element);

    internal OdfStyle EnsureStyle() => Document.Styles.EnsureAutomaticStyle(Element, OdfNamespaces.Text + "style-name", _family, _family == OdfStyleFamily.Paragraph ? "ofPr" : "ofRun", PartPath);
    internal void Dirty() => Document.MarkPartDirty(PartPath);
    internal T? Resolve<T>(Func<OdfStyle, T?> selector) where T : struct => Styles.Select(selector).FirstOrDefault(v => v.HasValue);
    internal string PartPath => Document.GetPartPath(Graphic);
    internal virtual IEnumerable<OdfStyle> Styles => OdfTextStyleResolver.Resolve(Document.Styles, Element, Graphic, PartPath);
}

/// <summary>A native styled inline span.</summary>
public sealed class OdfTextRun : OdfTextContent {
    internal OdfTextRun(OdfDocument document, XElement element, XElement graphic) : base(document, element, graphic, OdfStyleFamily.Text) { }
}
