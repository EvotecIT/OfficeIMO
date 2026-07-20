namespace OfficeIMO.Pdf;

/// <summary>
/// Inline text segment with basic styling.
/// </summary>
public sealed class TextRun {
    /// <summary>Text content of this run.</summary>
    public string Text { get; }
    /// <summary>True when bold style is applied.</summary>
    public bool Bold { get; }
    /// <summary>True when underline is applied.</summary>
    public bool Underline { get; }
    /// <summary>True when strikethrough is applied.</summary>
    public bool Strike { get; }
    /// <summary>True when italic style is applied.</summary>
    public bool Italic { get; }
    /// <summary>Run foreground color (if any).</summary>
    public PdfColor? Color { get; }
    /// <summary>Optional run background color, useful for highlights.</summary>
    public PdfColor? BackgroundColor { get; }
    /// <summary>Optional font size for this run. When null, the paragraph font size is used.</summary>
    public double? FontSize { get; }
    /// <summary>Optional standard PDF font for this run. When null, the paragraph/document font is used.</summary>
    public PdfStandardFont? Font { get; }
    /// <summary>Optional registered embedded font family for this run. This does not consume a standard-font compatibility slot; <see cref="Font"/> remains the fallback when the family is unavailable.</summary>
    public string? FontFamily { get; }
    /// <summary>Optional hyperlink URI associated with this run.</summary>
    public string? LinkUri { get; }
    /// <summary>Optional named destination associated with this run.</summary>
    public string? LinkDestinationName { get; }
    /// <summary>Optional hyperlink annotation contents, used by readers as link metadata.</summary>
    public string? LinkContents { get; }
    /// <summary>Baseline placement for this run.</summary>
    public PdfTextBaseline Baseline { get; }
    /// <summary>Leader fill used when this run represents a paragraph tab.</summary>
    public PdfTabLeaderStyle TabLeader { get; }
    /// <summary>Alignment used when this run represents a paragraph tab.</summary>
    public PdfTabAlignment TabAlignment { get; }
    /// <summary>Optional fixed-size visual carried by this run instead of text.</summary>
    public PdfInlineElement? InlineElement { get; }

    /// <summary>Create a new run with the specified styles.</summary>
    /// <param name="text">Run text.</param>
    /// <param name="bold">Whether to render bold.</param>
    /// <param name="underline">Whether to underline.</param>
    /// <param name="color">Run color or null to use defaults.</param>
    /// <param name="italic">Whether to render italic.</param>
    /// <param name="strike">Whether to render strikethrough.</param>
    /// <param name="fontSize">Optional run font size in points.</param>
    /// <param name="font">Optional standard PDF font for this run.</param>
    /// <param name="linkUri">Optional absolute URI or catalog-base-relative URI for link annotation.</param>
    /// <param name="linkContents">Optional link annotation contents; defaults to the run text when omitted.</param>
    /// <param name="baseline">Baseline placement for this run.</param>
    /// <param name="linkDestinationName">Optional named destination for an internal document link annotation.</param>
    /// <param name="tabLeader">Leader fill to render when the run text is a tab character.</param>
    /// <param name="tabAlignment">Alignment to use when the run text is a tab character.</param>
    /// <param name="backgroundColor">Optional run background color.</param>
    /// <param name="fontFamily">Optional registered embedded family name. <paramref name="font"/> is used as its fallback.</param>
    public TextRun(string text, bool bold = false, bool underline = false, PdfColor? color = null, bool italic = false, bool strike = false, double? fontSize = null, PdfStandardFont? font = null, string? linkUri = null, string? linkContents = null, PdfTextBaseline baseline = PdfTextBaseline.Normal, string? linkDestinationName = null, PdfTabLeaderStyle tabLeader = PdfTabLeaderStyle.None, PdfTabAlignment tabAlignment = PdfTabAlignment.Left, PdfColor? backgroundColor = null, string? fontFamily = null) {
        Guard.NotNull(text, nameof(text));
        Guard.TextBaseline(baseline, nameof(baseline));
        Guard.TabLeaderStyle(tabLeader, nameof(tabLeader));
        Guard.TabAlignment(tabAlignment, nameof(tabAlignment));
        if (fontSize.HasValue) {
            Guard.Positive(fontSize.Value, nameof(fontSize));
        }
        if (font.HasValue) {
            Guard.StandardFont(font.Value, nameof(font), "Text run font must be one of the supported standard PDF fonts.");
        }
        if (fontFamily != null) {
            Guard.NotNullOrWhiteSpace(fontFamily, nameof(fontFamily));
        }
        if (linkUri != null && linkDestinationName != null) {
            throw new System.ArgumentException("A text run link can target either a URI or a bookmark, not both.", nameof(linkDestinationName));
        }

        if ((tabLeader != PdfTabLeaderStyle.None || tabAlignment != PdfTabAlignment.Left) && text != "\t") {
            throw new System.ArgumentException("Tab leaders and alignment can only be applied to explicit tab runs.", nameof(tabAlignment));
        }

        bool hasLinkTarget = linkUri != null || linkDestinationName != null;
        if (linkContents != null && !hasLinkTarget) {
            throw new System.ArgumentException("Link annotation contents require a link target.", nameof(linkContents));
        }

        if (linkUri != null) {
            Guard.NotNullOrWhiteSpace(text, nameof(text));
            Guard.UriAction(linkUri, nameof(linkUri));
        }

        if (linkDestinationName != null) {
            Guard.NotNullOrWhiteSpace(text, nameof(text));
            Guard.NotNullOrWhiteSpace(linkDestinationName, nameof(linkDestinationName));
        }

        if (hasLinkTarget && linkContents != null) {
            Guard.NotNullOrWhiteSpace(linkContents, nameof(linkContents));
        }

        Text = text;
        Bold = bold;
        Underline = underline;
        Italic = italic;
        Strike = strike;
        Color = color;
        BackgroundColor = backgroundColor;
        FontSize = fontSize;
        Font = font;
        FontFamily = fontFamily?.Trim();
        LinkUri = linkUri;
        LinkDestinationName = linkDestinationName;
        LinkContents = hasLinkTarget ? linkContents ?? text : null;
        Baseline = baseline;
        TabLeader = tabLeader;
        TabAlignment = tabAlignment;
        InlineElement = null;
    }

    private TextRun(PdfInlineElement inlineElement)
        : this(string.Empty) {
        InlineElement = inlineElement ?? throw new ArgumentNullException(nameof(inlineElement));
    }

    /// <summary>Create a normal (unstyled) run.</summary>
    public static TextRun Normal(string text, PdfColor? color = null, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) => new TextRun(text, bold: false, underline: false, color: color, italic: false, strike: false, fontSize: fontSize, font: font, backgroundColor: backgroundColor, fontFamily: fontFamily);
    /// <summary>Create an explicit line-break run.</summary>
    public static TextRun LineBreak() => new TextRun("\n", bold: false, underline: false, color: null, italic: false, strike: false);
    /// <summary>Create an explicit paragraph tab run.</summary>
    public static TextRun Tab(PdfTabLeaderStyle leader = PdfTabLeaderStyle.None, PdfTabAlignment alignment = PdfTabAlignment.Left) => new TextRun("\t", tabLeader: leader, tabAlignment: alignment);
    /// <summary>Create a fixed-size inline visual run.</summary>
    public static TextRun Inline(PdfInlineElement element) => new TextRun(element);
    /// <summary>Create a bold run.</summary>
    public static TextRun Bolded(string text, PdfColor? color = null, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) => new TextRun(text, bold: true, underline: false, color: color, italic: false, strike: false, fontSize: fontSize, font: font, backgroundColor: backgroundColor, fontFamily: fontFamily);
    /// <summary>Create an underlined run.</summary>
    public static TextRun Underlined(string text, PdfColor? color = null, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) => new TextRun(text, bold: false, underline: true, color: color, italic: false, strike: false, fontSize: fontSize, font: font, backgroundColor: backgroundColor, fontFamily: fontFamily);
    /// <summary>Create an italic run.</summary>
    public static TextRun Italicized(string text, PdfColor? color = null, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) => new TextRun(text, bold: false, underline: false, color: color, italic: true, strike: false, fontSize: fontSize, font: font, backgroundColor: backgroundColor, fontFamily: fontFamily);
    /// <summary>Create a bold and underlined run.</summary>
    public static TextRun BoldUnderlined(string text, PdfColor? color = null, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) => new TextRun(text, bold: true, underline: true, color: color, italic: false, strike: false, fontSize: fontSize, font: font, backgroundColor: backgroundColor, fontFamily: fontFamily);
    /// <summary>Create a bold and italic run.</summary>
    public static TextRun BoldItalic(string text, PdfColor? color = null, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) => new TextRun(text, bold: true, underline: false, color: color, italic: true, strike: false, fontSize: fontSize, font: font, backgroundColor: backgroundColor, fontFamily: fontFamily);
    /// <summary>Create a strikethrough run.</summary>
    public static TextRun Strikethrough(string text, PdfColor? color = null, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) => new TextRun(text, bold: false, underline: false, color: color, italic: false, strike: true, fontSize: fontSize, font: font, backgroundColor: backgroundColor, fontFamily: fontFamily);
    /// <summary>Create a superscript run.</summary>
    public static TextRun Superscript(string text, PdfColor? color = null, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) => new TextRun(text, bold: false, underline: false, color: color, italic: false, strike: false, fontSize: fontSize, font: font, baseline: PdfTextBaseline.Superscript, backgroundColor: backgroundColor, fontFamily: fontFamily);
    /// <summary>Create a subscript run.</summary>
    public static TextRun Subscript(string text, PdfColor? color = null, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) => new TextRun(text, bold: false, underline: false, color: color, italic: false, strike: false, fontSize: fontSize, font: font, baseline: PdfTextBaseline.Subscript, backgroundColor: backgroundColor, fontFamily: fontFamily);
    /// <summary>Create a hyperlink run that points to a URI.</summary>
    /// <param name="text">Link text.</param>
    /// <param name="uri">Absolute URI or catalog-base-relative URI.</param>
    /// <param name="color">Optional link color.</param>
    /// <param name="underline">Whether to underline the link text.</param>
    /// <param name="contents">Optional link annotation contents.</param>
    /// <param name="baseline">Baseline placement for this run.</param>
    /// <param name="fontSize">Optional run font size in points.</param>
    /// <param name="backgroundColor">Optional run background color.</param>
    /// <param name="font">Optional standard font family for this run.</param>
    /// <param name="fontFamily">Optional registered embedded family name. <paramref name="font"/> remains its fallback.</param>
    public static TextRun Link(string text, string uri, PdfColor? color = null, bool underline = true, string? contents = null, PdfTextBaseline baseline = PdfTextBaseline.Normal, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) {
        Guard.UriAction(uri, nameof(uri));
        return new TextRun(text, bold: false, underline: underline, color: color, italic: false, strike: false, fontSize: fontSize, font: font, linkUri: uri, linkContents: contents, baseline: baseline, backgroundColor: backgroundColor, fontFamily: fontFamily);
    }
    /// <summary>Create a hyperlink run that points to a document bookmark.</summary>
    /// <param name="text">Link text.</param>
    /// <param name="bookmarkName">Named destination created with <see cref="PdfDocument.Bookmark(string)"/>.</param>
    /// <param name="color">Optional link color.</param>
    /// <param name="underline">Whether to underline the link text.</param>
    /// <param name="contents">Optional link annotation contents.</param>
    /// <param name="baseline">Baseline placement for this run.</param>
    /// <param name="fontSize">Optional run font size in points.</param>
    /// <param name="backgroundColor">Optional run background color.</param>
    /// <param name="font">Optional standard font family for this run.</param>
    /// <param name="fontFamily">Optional registered embedded family name. <paramref name="font"/> remains its fallback.</param>
    public static TextRun LinkToBookmark(string text, string bookmarkName, PdfColor? color = null, bool underline = true, string? contents = null, PdfTextBaseline baseline = PdfTextBaseline.Normal, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) {
        Guard.NotNullOrWhiteSpace(bookmarkName, nameof(bookmarkName));
        return new TextRun(text, bold: false, underline: underline, color: color, italic: false, strike: false, fontSize: fontSize, font: font, linkContents: contents, baseline: baseline, linkDestinationName: bookmarkName, backgroundColor: backgroundColor, fontFamily: fontFamily);
    }
}
