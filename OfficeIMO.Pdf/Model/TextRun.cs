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

    /// <summary>Create a new run with the specified styles.</summary>
    /// <param name="text">Run text.</param>
    /// <param name="bold">Whether to render bold.</param>
    /// <param name="underline">Whether to underline.</param>
    /// <param name="color">Run color or null to use defaults.</param>
    /// <param name="italic">Whether to render italic.</param>
    /// <param name="strike">Whether to render strikethrough.</param>
    /// <param name="linkUri">Optional absolute URI for link annotation.</param>
    /// <param name="linkContents">Optional link annotation contents; defaults to the run text when omitted.</param>
    /// <param name="baseline">Baseline placement for this run.</param>
    /// <param name="linkDestinationName">Optional named destination for an internal document link annotation.</param>
    /// <param name="tabLeader">Leader fill to render when the run text is a tab character.</param>
    /// <param name="tabAlignment">Alignment to use when the run text is a tab character.</param>
    public TextRun(string text, bool bold = false, bool underline = false, PdfColor? color = null, bool italic = false, bool strike = false, string? linkUri = null, string? linkContents = null, PdfTextBaseline baseline = PdfTextBaseline.Normal, string? linkDestinationName = null, PdfTabLeaderStyle tabLeader = PdfTabLeaderStyle.None, PdfTabAlignment tabAlignment = PdfTabAlignment.Left) {
        Guard.NotNull(text, nameof(text));
        Guard.TextBaseline(baseline, nameof(baseline));
        Guard.TabLeaderStyle(tabLeader, nameof(tabLeader));
        Guard.TabAlignment(tabAlignment, nameof(tabAlignment));
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
            Guard.AbsoluteUri(linkUri, nameof(linkUri));
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
        LinkUri = linkUri;
        LinkDestinationName = linkDestinationName;
        LinkContents = hasLinkTarget ? linkContents ?? text : null;
        Baseline = baseline;
        TabLeader = tabLeader;
        TabAlignment = tabAlignment;
    }

    /// <summary>Create a normal (unstyled) run.</summary>
    public static TextRun Normal(string text, PdfColor? color = null) => new TextRun(text, bold: false, underline: false, color: color, italic: false, strike: false);
    /// <summary>Create an explicit line-break run.</summary>
    public static TextRun LineBreak() => new TextRun("\n", bold: false, underline: false, color: null, italic: false, strike: false);
    /// <summary>Create an explicit paragraph tab run.</summary>
    public static TextRun Tab(PdfTabLeaderStyle leader = PdfTabLeaderStyle.None, PdfTabAlignment alignment = PdfTabAlignment.Left) => new TextRun("\t", tabLeader: leader, tabAlignment: alignment);
    /// <summary>Create a bold run.</summary>
    public static TextRun Bolded(string text, PdfColor? color = null) => new TextRun(text, bold: true, underline: false, color: color, italic: false, strike: false);
    /// <summary>Create an underlined run.</summary>
    public static TextRun Underlined(string text, PdfColor? color = null) => new TextRun(text, bold: false, underline: true, color: color, italic: false, strike: false);
    /// <summary>Create an italic run.</summary>
    public static TextRun Italicized(string text, PdfColor? color = null) => new TextRun(text, bold: false, underline: false, color: color, italic: true, strike: false);
    /// <summary>Create a bold and underlined run.</summary>
    public static TextRun BoldUnderlined(string text, PdfColor? color = null) => new TextRun(text, bold: true, underline: true, color: color, italic: false, strike: false);
    /// <summary>Create a bold and italic run.</summary>
    public static TextRun BoldItalic(string text, PdfColor? color = null) => new TextRun(text, bold: true, underline: false, color: color, italic: true, strike: false);
    /// <summary>Create a strikethrough run.</summary>
    public static TextRun Strikethrough(string text, PdfColor? color = null) => new TextRun(text, bold: false, underline: false, color: color, italic: false, strike: true);
    /// <summary>Create a superscript run.</summary>
    public static TextRun Superscript(string text, PdfColor? color = null) => new TextRun(text, bold: false, underline: false, color: color, italic: false, strike: false, baseline: PdfTextBaseline.Superscript);
    /// <summary>Create a subscript run.</summary>
    public static TextRun Subscript(string text, PdfColor? color = null) => new TextRun(text, bold: false, underline: false, color: color, italic: false, strike: false, baseline: PdfTextBaseline.Subscript);
    /// <summary>Create a hyperlink run that points to a URI.</summary>
    /// <param name="text">Link text.</param>
    /// <param name="uri">Absolute URI.</param>
    /// <param name="color">Optional link color.</param>
    /// <param name="underline">Whether to underline the link text.</param>
    /// <param name="contents">Optional link annotation contents.</param>
    /// <param name="baseline">Baseline placement for this run.</param>
    public static TextRun Link(string text, string uri, PdfColor? color = null, bool underline = true, string? contents = null, PdfTextBaseline baseline = PdfTextBaseline.Normal) {
        Guard.AbsoluteUri(uri, nameof(uri));
        return new TextRun(text, bold: false, underline: underline, color: color, italic: false, strike: false, linkUri: uri, linkContents: contents, baseline: baseline);
    }
    /// <summary>Create a hyperlink run that points to a document bookmark.</summary>
    /// <param name="text">Link text.</param>
    /// <param name="bookmarkName">Named destination created with <see cref="PdfDoc.Bookmark(string)"/>.</param>
    /// <param name="color">Optional link color.</param>
    /// <param name="underline">Whether to underline the link text.</param>
    /// <param name="contents">Optional link annotation contents.</param>
    /// <param name="baseline">Baseline placement for this run.</param>
    public static TextRun LinkToBookmark(string text, string bookmarkName, PdfColor? color = null, bool underline = true, string? contents = null, PdfTextBaseline baseline = PdfTextBaseline.Normal) {
        Guard.NotNullOrWhiteSpace(bookmarkName, nameof(bookmarkName));
        return new TextRun(text, bold: false, underline: underline, color: color, italic: false, strike: false, linkContents: contents, baseline: baseline, linkDestinationName: bookmarkName);
    }
}
