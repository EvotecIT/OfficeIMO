namespace OfficeIMO.Pdf;

/// <summary>
/// Inline text segment with basic styling.
/// </summary>
public sealed class PdfTextRun {
    /// <summary>Text content of this run.</summary>
    public string Text { get; }
    /// <summary>True when bold style is applied.</summary>
    public bool Bold { get; }
    /// <summary>True when underline is applied.</summary>
    public bool Underline => UnderlineStyle != OfficeIMO.Drawing.OfficeTextDecorationStyle.None;
    /// <summary>Underline pattern rendered by the PDF writer.</summary>
    public OfficeIMO.Drawing.OfficeTextDecorationStyle UnderlineStyle { get; }
    /// <summary>True when strikethrough is applied.</summary>
    public bool Strike => StrikeStyle != OfficeIMO.Drawing.OfficeTextDecorationStyle.None;
    /// <summary>Strikethrough pattern rendered by the PDF writer.</summary>
    public OfficeIMO.Drawing.OfficeTextDecorationStyle StrikeStyle { get; }
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
    /// <param name="backgroundColor">Optional run background color.</param>
    /// <param name="fontFamily">Optional registered embedded family name. <paramref name="font"/> is used as its fallback.</param>
    public PdfTextRun(string text, bool bold, bool underline, PdfColor? color, bool italic, bool strike, double? fontSize, PdfStandardFont? font, string? linkUri, string? linkContents, PdfTextBaseline baseline, string? linkDestinationName, PdfTabLeaderStyle tabLeader, PdfColor? backgroundColor, string? fontFamily)
        : this(text, bold, underline, color, italic, strike, fontSize, font, linkUri, linkContents, baseline, linkDestinationName, tabLeader, PdfTabAlignment.Left, backgroundColor, fontFamily) {
    }

    /// <summary>Create a run using the pre-typography constructor signature with tab alignment.</summary>
    public PdfTextRun(string text, bool bold, bool underline, PdfColor? color, bool italic, bool strike, double? fontSize, PdfStandardFont? font, string? linkUri, string? linkContents, PdfTextBaseline baseline, string? linkDestinationName, PdfTabLeaderStyle tabLeader, PdfTabAlignment tabAlignment, PdfColor? backgroundColor, string? fontFamily)
        : this(text, bold, underline, color, italic, strike, fontSize, font, linkUri, linkContents, baseline, linkDestinationName, tabLeader, tabAlignment, backgroundColor, fontFamily,
            OfficeIMO.Drawing.OfficeTextDecorationStyle.None, OfficeIMO.Drawing.OfficeTextDecorationStyle.None) {
    }

    /// <summary>Create a new run with the specified styles and tab alignment.</summary>
    public PdfTextRun(string text, bool bold = false, bool underline = false, PdfColor? color = null, bool italic = false, bool strike = false, double? fontSize = null, PdfStandardFont? font = null, string? linkUri = null, string? linkContents = null, PdfTextBaseline baseline = PdfTextBaseline.Normal, string? linkDestinationName = null, PdfTabLeaderStyle tabLeader = PdfTabLeaderStyle.None, PdfTabAlignment tabAlignment = PdfTabAlignment.Left, PdfColor? backgroundColor = null, string? fontFamily = null, OfficeIMO.Drawing.OfficeTextDecorationStyle underlineStyle = OfficeIMO.Drawing.OfficeTextDecorationStyle.None, OfficeIMO.Drawing.OfficeTextDecorationStyle strikeStyle = OfficeIMO.Drawing.OfficeTextDecorationStyle.None) {
        Guard.NotNull(text, nameof(text));
        Guard.TextBaseline(baseline, nameof(baseline));
        Guard.TabLeaderStyle(tabLeader, nameof(tabLeader));
        Guard.TabAlignment(tabAlignment, nameof(tabAlignment));
        if (underlineStyle < OfficeIMO.Drawing.OfficeTextDecorationStyle.None || underlineStyle > OfficeIMO.Drawing.OfficeTextDecorationStyle.Wavy) {
            throw new System.ArgumentOutOfRangeException(nameof(underlineStyle));
        }
        if (strikeStyle < OfficeIMO.Drawing.OfficeTextDecorationStyle.None || strikeStyle > OfficeIMO.Drawing.OfficeTextDecorationStyle.Wavy) {
            throw new System.ArgumentOutOfRangeException(nameof(strikeStyle));
        }
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
        UnderlineStyle = underlineStyle != OfficeIMO.Drawing.OfficeTextDecorationStyle.None
            ? underlineStyle
            : underline ? OfficeIMO.Drawing.OfficeTextDecorationStyle.Single : OfficeIMO.Drawing.OfficeTextDecorationStyle.None;
        Italic = italic;
        StrikeStyle = strikeStyle != OfficeIMO.Drawing.OfficeTextDecorationStyle.None
            ? strikeStyle
            : strike ? OfficeIMO.Drawing.OfficeTextDecorationStyle.Single : OfficeIMO.Drawing.OfficeTextDecorationStyle.None;
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

    private PdfTextRun(PdfInlineElement inlineElement)
        : this(string.Empty) {
        InlineElement = inlineElement ?? throw new ArgumentNullException(nameof(inlineElement));
    }

    /// <summary>Create a normal (unstyled) run.</summary>
    public static PdfTextRun Normal(string text, PdfColor? color = null, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) => new PdfTextRun(text, bold: false, underline: false, color: color, italic: false, strike: false, fontSize: fontSize, font: font, backgroundColor: backgroundColor, fontFamily: fontFamily);
    /// <summary>Create an explicit line-break run.</summary>
    public static PdfTextRun LineBreak() => new PdfTextRun("\n", bold: false, underline: false, color: null, italic: false, strike: false);
    /// <summary>Create an explicit paragraph tab run.</summary>
    public static PdfTextRun Tab(PdfTabLeaderStyle leader) => Tab(leader, PdfTabAlignment.Left);
    /// <summary>Create an explicit paragraph tab run with alignment.</summary>
    public static PdfTextRun Tab(PdfTabLeaderStyle leader = PdfTabLeaderStyle.None, PdfTabAlignment alignment = PdfTabAlignment.Left) => new PdfTextRun("\t", tabLeader: leader, tabAlignment: alignment);
    /// <summary>Create a fixed-size inline visual run.</summary>
    public static PdfTextRun Inline(PdfInlineElement element) => new PdfTextRun(element);
    /// <summary>Create a bold run.</summary>
    public static PdfTextRun Bolded(string text, PdfColor? color = null, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) => new PdfTextRun(text, bold: true, underline: false, color: color, italic: false, strike: false, fontSize: fontSize, font: font, backgroundColor: backgroundColor, fontFamily: fontFamily);
    /// <summary>Create an underlined run.</summary>
    public static PdfTextRun Underlined(string text, PdfColor? color = null, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) => new PdfTextRun(text, bold: false, underline: true, color: color, italic: false, strike: false, fontSize: fontSize, font: font, backgroundColor: backgroundColor, fontFamily: fontFamily);
    /// <summary>Create an italic run.</summary>
    public static PdfTextRun Italicized(string text, PdfColor? color = null, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) => new PdfTextRun(text, bold: false, underline: false, color: color, italic: true, strike: false, fontSize: fontSize, font: font, backgroundColor: backgroundColor, fontFamily: fontFamily);
    /// <summary>Create a bold and underlined run.</summary>
    public static PdfTextRun BoldUnderlined(string text, PdfColor? color = null, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) => new PdfTextRun(text, bold: true, underline: true, color: color, italic: false, strike: false, fontSize: fontSize, font: font, backgroundColor: backgroundColor, fontFamily: fontFamily);
    /// <summary>Create a bold and italic run.</summary>
    public static PdfTextRun BoldItalic(string text, PdfColor? color = null, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) => new PdfTextRun(text, bold: true, underline: false, color: color, italic: true, strike: false, fontSize: fontSize, font: font, backgroundColor: backgroundColor, fontFamily: fontFamily);
    /// <summary>Create a strikethrough run.</summary>
    public static PdfTextRun Strikethrough(string text, PdfColor? color = null, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) => new PdfTextRun(text, bold: false, underline: false, color: color, italic: false, strike: true, fontSize: fontSize, font: font, backgroundColor: backgroundColor, fontFamily: fontFamily);
    /// <summary>Create a superscript run.</summary>
    public static PdfTextRun Superscript(string text, PdfColor? color = null, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) => new PdfTextRun(text, bold: false, underline: false, color: color, italic: false, strike: false, fontSize: fontSize, font: font, baseline: PdfTextBaseline.Superscript, backgroundColor: backgroundColor, fontFamily: fontFamily);
    /// <summary>Create a subscript run.</summary>
    public static PdfTextRun Subscript(string text, PdfColor? color = null, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null) => new PdfTextRun(text, bold: false, underline: false, color: color, italic: false, strike: false, fontSize: fontSize, font: font, baseline: PdfTextBaseline.Subscript, backgroundColor: backgroundColor, fontFamily: fontFamily);
    /// <summary>Create a copy with transformed text casing while preserving all run formatting and link metadata.</summary>
    public PdfTextRun WithTextCase(OfficeIMO.Drawing.OfficeTextCase textCase, System.Globalization.CultureInfo? culture = null) {
        if (InlineElement != null) return this;
        string transformed = OfficeIMO.Drawing.OfficeTextCaseTransformer.Apply(Text, textCase, culture);
        return new PdfTextRun(transformed, Bold, Underline, Color, Italic, Strike, FontSize, Font, LinkUri, LinkContents, Baseline, LinkDestinationName, TabLeader, TabAlignment, BackgroundColor, FontFamily, UnderlineStyle, StrikeStyle);
    }
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
    /// <param name="underlineStyle">Optional underline pattern.</param>
    /// <param name="strikeStyle">Optional strikethrough pattern.</param>
    public static PdfTextRun Link(string text, string uri, PdfColor? color = null, bool underline = true, string? contents = null, PdfTextBaseline baseline = PdfTextBaseline.Normal, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null, OfficeIMO.Drawing.OfficeTextDecorationStyle underlineStyle = OfficeIMO.Drawing.OfficeTextDecorationStyle.None, OfficeIMO.Drawing.OfficeTextDecorationStyle strikeStyle = OfficeIMO.Drawing.OfficeTextDecorationStyle.None) {
        Guard.UriAction(uri, nameof(uri));
        return new PdfTextRun(text, bold: false, underline: underline, color: color, italic: false, strike: false, fontSize: fontSize, font: font, linkUri: uri, linkContents: contents, baseline: baseline, backgroundColor: backgroundColor, fontFamily: fontFamily, underlineStyle: underlineStyle, strikeStyle: strikeStyle);
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
    /// <param name="underlineStyle">Optional underline pattern.</param>
    /// <param name="strikeStyle">Optional strikethrough pattern.</param>
    public static PdfTextRun LinkToBookmark(string text, string bookmarkName, PdfColor? color = null, bool underline = true, string? contents = null, PdfTextBaseline baseline = PdfTextBaseline.Normal, double? fontSize = null, PdfColor? backgroundColor = null, PdfStandardFont? font = null, string? fontFamily = null, OfficeIMO.Drawing.OfficeTextDecorationStyle underlineStyle = OfficeIMO.Drawing.OfficeTextDecorationStyle.None, OfficeIMO.Drawing.OfficeTextDecorationStyle strikeStyle = OfficeIMO.Drawing.OfficeTextDecorationStyle.None) {
        Guard.NotNullOrWhiteSpace(bookmarkName, nameof(bookmarkName));
        return new PdfTextRun(text, bold: false, underline: underline, color: color, italic: false, strike: false, fontSize: fontSize, font: font, linkContents: contents, baseline: baseline, linkDestinationName: bookmarkName, backgroundColor: backgroundColor, fontFamily: fontFamily, underlineStyle: underlineStyle, strikeStyle: strikeStyle);
    }
}
