namespace OfficeIMO.Pdf;

/// <summary>
/// Fluent builder for rich paragraphs made of styled text runs.
/// </summary>
public sealed class PdfParagraphBuilder {
    private readonly System.Collections.Generic.List<PdfTextRun> _runs = new();
    private PdfColor? _currentColor;
    private bool _currentBold;
    private bool _currentItalic;
    private bool _currentUnderline;
    private bool _currentStrike;
    private OfficeIMO.Drawing.OfficeTextDecorationStyle _currentUnderlineStyle;
    private OfficeIMO.Drawing.OfficeTextDecorationStyle _currentStrikeStyle;
    private PdfTextBaseline _currentBaseline;
    private double? _currentFontSize;
    private PdfStandardFont? _currentFont;
    private string? _currentFontFamily;
    private PdfColor? _currentBackgroundColor;

    /// <summary>Paragraph alignment.</summary>
    public PdfAlign Align { get; }
    /// <summary>Default text color applied when no run color is specified.</summary>
    public PdfColor? DefaultColor { get; }

    /// <summary>Create a new paragraph builder.</summary>
    /// <param name="align">Paragraph alignment.</param>
    /// <param name="defaultColor">Optional default text color.</param>
    public PdfParagraphBuilder(PdfAlign align, PdfColor? defaultColor) {
        Guard.ParagraphAlign(align, nameof(align), "Paragraph");
        Align = align; DefaultColor = defaultColor; _currentColor = defaultColor;
    }

    /// <summary>Sets the current run color.</summary>
    public PdfParagraphBuilder Color(PdfColor color) { _currentColor = color; return this; }
    /// <summary>Resets the current run color to the paragraph default color.</summary>
    public PdfParagraphBuilder ResetColor() { _currentColor = DefaultColor; return this; }
    /// <summary>Sets the current run font size in points.</summary>
    public PdfParagraphBuilder FontSize(double fontSize) { Guard.Positive(fontSize, nameof(fontSize)); _currentFontSize = fontSize; return this; }
    /// <summary>Resets the current run font size to the paragraph default font size.</summary>
    public PdfParagraphBuilder ResetFontSize() { _currentFontSize = null; return this; }
    /// <summary>Sets the current standard PDF font for subsequent runs.</summary>
    public PdfParagraphBuilder Font(PdfStandardFont font) { Guard.StandardFont(font, nameof(font), "Text run font must be one of the supported standard PDF fonts."); _currentFont = font; _currentFontFamily = null; return this; }
    /// <summary>Sets a registered embedded named family for subsequent runs.</summary>
    public PdfParagraphBuilder FontFamily(string familyName) { Guard.NotNullOrWhiteSpace(familyName, nameof(familyName)); _currentFontFamily = familyName.Trim(); _currentFont = null; return this; }
    /// <summary>Resets the current run font to the paragraph/document font.</summary>
    public PdfParagraphBuilder ResetFont() { _currentFont = null; _currentFontFamily = null; return this; }
    /// <summary>Sets the current run background color.</summary>
    public PdfParagraphBuilder BackgroundColor(PdfColor color) { _currentBackgroundColor = color; return this; }
    /// <summary>Resets the current run background color.</summary>
    public PdfParagraphBuilder ResetBackgroundColor() { _currentBackgroundColor = null; return this; }
    /// <summary>Enables or disables bold for subsequent runs.</summary>
    public PdfParagraphBuilder Bold(bool enable = true) { _currentBold = enable; return this; }
    /// <summary>Enables or disables italic for subsequent runs.</summary>
    public PdfParagraphBuilder Italic(bool enable = true) { _currentItalic = enable; return this; }
    /// <summary>Enables or disables underline for subsequent runs.</summary>
    public PdfParagraphBuilder Underline(bool enable = true) { _currentUnderline = enable; _currentUnderlineStyle = enable ? OfficeIMO.Drawing.OfficeTextDecorationStyle.Single : OfficeIMO.Drawing.OfficeTextDecorationStyle.None; return this; }
    /// <summary>Sets the underline pattern for subsequent runs.</summary>
    public PdfParagraphBuilder Underline(OfficeIMO.Drawing.OfficeTextDecorationStyle style) { ValidateDecorationStyle(style, nameof(style)); _currentUnderlineStyle = style; _currentUnderline = style != OfficeIMO.Drawing.OfficeTextDecorationStyle.None; return this; }
    /// <summary>Enables or disables strikethrough for subsequent runs.</summary>
    public PdfParagraphBuilder Strike(bool enable = true) { _currentStrike = enable; _currentStrikeStyle = enable ? OfficeIMO.Drawing.OfficeTextDecorationStyle.Single : OfficeIMO.Drawing.OfficeTextDecorationStyle.None; return this; }
    /// <summary>Sets the strikethrough pattern for subsequent runs.</summary>
    public PdfParagraphBuilder Strike(OfficeIMO.Drawing.OfficeTextDecorationStyle style) { ValidateDecorationStyle(style, nameof(style)); _currentStrikeStyle = style; _currentStrike = style != OfficeIMO.Drawing.OfficeTextDecorationStyle.None; return this; }
    /// <summary>Enables or disables superscript baseline placement for subsequent runs.</summary>
    public PdfParagraphBuilder Superscript(bool enable = true) { _currentBaseline = enable ? PdfTextBaseline.Superscript : PdfTextBaseline.Normal; return this; }
    /// <summary>Enables or disables subscript baseline placement for subsequent runs.</summary>
    public PdfParagraphBuilder Subscript(bool enable = true) { _currentBaseline = enable ? PdfTextBaseline.Subscript : PdfTextBaseline.Normal; return this; }
    /// <summary>Sets the current baseline placement for subsequent runs.</summary>
    public PdfParagraphBuilder Baseline(PdfTextBaseline baseline) { Guard.TextBaseline(baseline, nameof(baseline)); _currentBaseline = baseline; return this; }

    /// <summary>Adds a text run using the current style flags.</summary>
    public PdfParagraphBuilder Text(string text) { _runs.Add(new PdfTextRun(text, _currentBold, _currentUnderline, _currentColor, _currentItalic, _currentStrike, fontSize: _currentFontSize, font: _currentFont, baseline: _currentBaseline, backgroundColor: _currentBackgroundColor, fontFamily: _currentFontFamily, underlineStyle: _currentUnderlineStyle, strikeStyle: _currentStrikeStyle)); return this; }
    /// <summary>Adds text planned through an embedded-font fallback set while preserving the current run styling.</summary>
    public PdfParagraphBuilder FallbackText(PdfEmbeddedFontFallbackSet fallbackSet, string text, string source = "") {
        Guard.NotNull(fallbackSet, nameof(fallbackSet));
        Guard.NotNull(text, nameof(text));
        _runs.AddRange(fallbackSet.PlanTextRuns(text, source, CreateCurrentStyleTemplate()));
        return this;
    }
    /// <summary>Adds already prepared rich text runs, preserving their per-run styling and font selections.</summary>
    public PdfParagraphBuilder Runs(System.Collections.Generic.IEnumerable<PdfTextRun> runs) {
        Guard.NotNull(runs, nameof(runs));
        _runs.AddRange(runs);
        return this;
    }
    /// <summary>Adds an explicit line break inside the current paragraph.</summary>
    public PdfParagraphBuilder LineBreak() { _runs.Add(PdfTextRun.LineBreak()); return this; }
    /// <summary>Adds an explicit paragraph tab using the current style flags.</summary>
    public PdfParagraphBuilder Tab(PdfTabLeaderStyle leader) => Tab(leader, PdfTabAlignment.Left);
    /// <summary>Adds an aligned paragraph tab using the current style flags.</summary>
    public PdfParagraphBuilder Tab(PdfTabLeaderStyle leader = PdfTabLeaderStyle.None, PdfTabAlignment alignment = PdfTabAlignment.Left) { _runs.Add(new PdfTextRun("\t", _currentBold, _currentUnderline, _currentColor, _currentItalic, _currentStrike, fontSize: _currentFontSize, font: _currentFont, baseline: _currentBaseline, tabLeader: leader, tabAlignment: alignment, backgroundColor: _currentBackgroundColor, fontFamily: _currentFontFamily, underlineStyle: _currentUnderlineStyle, strikeStyle: _currentStrikeStyle)); return this; }
    /// <summary>Adds a fixed-size visual that participates in paragraph wrapping.</summary>
    public PdfParagraphBuilder Inline(PdfInlineElement element) { Guard.NotNull(element, nameof(element)); _runs.Add(PdfTextRun.Inline(element)); return this; }
    /// <summary>Adds an image that participates in paragraph wrapping.</summary>
    public PdfParagraphBuilder InlineImage(byte[] imageBytes, double width, double height, string? alternativeText = null, OfficeIMO.Drawing.OfficeImageFit fit = OfficeIMO.Drawing.OfficeImageFit.Contain, double baselineOffset = 0D) =>
        Inline(new PdfInlineImage(imageBytes, width, height, alternativeText, fit, baselineOffset));
    /// <summary>Adds a filled and/or bordered box that participates in paragraph wrapping.</summary>
    public PdfParagraphBuilder InlineBox(double width, double height, PdfColor? background = null, PdfColor? borderColor = null, double borderWidth = 0.5D, string? alternativeText = null, double baselineOffset = 0D) =>
        Inline(new PdfInlineBox(width, height, background, borderColor, borderWidth, alternativeText, baselineOffset));
    /// <summary>Adds a bold text run.</summary>
    public PdfParagraphBuilder Bold(string text, PdfColor? color = null) { _runs.Add(new PdfTextRun(text, bold: true, underline: false, color: color ?? _currentColor, italic: false, fontSize: _currentFontSize, font: _currentFont, backgroundColor: _currentBackgroundColor, fontFamily: _currentFontFamily)); return this; }
    /// <summary>Adds an italic text run.</summary>
    public PdfParagraphBuilder Italic(string text, PdfColor? color = null) { _runs.Add(new PdfTextRun(text, bold: false, underline: false, color: color ?? _currentColor, italic: true, fontSize: _currentFontSize, font: _currentFont, backgroundColor: _currentBackgroundColor, fontFamily: _currentFontFamily)); return this; }
    /// <summary>Adds an underlined text run.</summary>
    public PdfParagraphBuilder Underlined(string text, PdfColor? color = null) { _runs.Add(new PdfTextRun(text, bold: false, underline: true, color: color ?? _currentColor, italic: false, fontSize: _currentFontSize, font: _currentFont, backgroundColor: _currentBackgroundColor, fontFamily: _currentFontFamily)); return this; }
    /// <summary>Adds a text run with the requested underline pattern.</summary>
    public PdfParagraphBuilder Underlined(string text, OfficeIMO.Drawing.OfficeTextDecorationStyle style, PdfColor? color = null) { _runs.Add(new PdfTextRun(text, color: color ?? _currentColor, fontSize: _currentFontSize, font: _currentFont, backgroundColor: _currentBackgroundColor, fontFamily: _currentFontFamily, underlineStyle: style)); return this; }
    /// <summary>Adds a strikethrough text run.</summary>
    public PdfParagraphBuilder Strikethrough(string text, PdfColor? color = null) { _runs.Add(new PdfTextRun(text, bold: false, underline: false, color: color ?? _currentColor, italic: false, strike: true, fontSize: _currentFontSize, font: _currentFont, backgroundColor: _currentBackgroundColor, fontFamily: _currentFontFamily)); return this; }
    /// <summary>Adds a text run with the requested strikethrough pattern.</summary>
    public PdfParagraphBuilder Strikethrough(string text, OfficeIMO.Drawing.OfficeTextDecorationStyle style, PdfColor? color = null) { _runs.Add(new PdfTextRun(text, color: color ?? _currentColor, fontSize: _currentFontSize, font: _currentFont, backgroundColor: _currentBackgroundColor, fontFamily: _currentFontFamily, strikeStyle: style)); return this; }
    /// <summary>Adds a superscript text run.</summary>
    public PdfParagraphBuilder Superscript(string text, PdfColor? color = null) { _runs.Add(new PdfTextRun(text, bold: false, underline: false, color: color ?? _currentColor, italic: false, strike: false, fontSize: _currentFontSize, font: _currentFont, baseline: PdfTextBaseline.Superscript, backgroundColor: _currentBackgroundColor, fontFamily: _currentFontFamily)); return this; }
    /// <summary>Adds a subscript text run.</summary>
    public PdfParagraphBuilder Subscript(string text, PdfColor? color = null) { _runs.Add(new PdfTextRun(text, bold: false, underline: false, color: color ?? _currentColor, italic: false, strike: false, fontSize: _currentFontSize, font: _currentFont, baseline: PdfTextBaseline.Subscript, backgroundColor: _currentBackgroundColor, fontFamily: _currentFontFamily)); return this; }
    /// <summary>Adds a hyperlink text run.</summary>
    /// <param name="text">Link text.</param>
    /// <param name="uri">Absolute URI or catalog-base-relative URI to open.</param>
    /// <param name="color">Optional link color.</param>
    /// <param name="underline">Whether to underline the link text (default true).</param>
    /// <param name="contents">Optional link annotation contents; defaults to the link text when omitted.</param>
    public PdfParagraphBuilder Link(string text, string uri, PdfColor? color = null, bool underline = true, string? contents = null) { _runs.Add(PdfTextRun.Link(text, uri, color ?? _currentColor, underline, contents, _currentBaseline, _currentFontSize, _currentBackgroundColor, _currentFont, _currentFontFamily, underline ? _currentUnderlineStyle : OfficeIMO.Drawing.OfficeTextDecorationStyle.None, _currentStrikeStyle)); return this; }
    /// <summary>Adds a hyperlink text run that points to a document bookmark.</summary>
    /// <param name="text">Link text.</param>
    /// <param name="bookmarkName">Named destination created with <see cref="PdfDocument.Bookmark(string)"/>.</param>
    /// <param name="color">Optional link color.</param>
    /// <param name="underline">Whether to underline the link text (default true).</param>
    /// <param name="contents">Optional link annotation contents; defaults to the link text when omitted.</param>
    public PdfParagraphBuilder LinkToBookmark(string text, string bookmarkName, PdfColor? color = null, bool underline = true, string? contents = null) { _runs.Add(PdfTextRun.LinkToBookmark(text, bookmarkName, color ?? _currentColor, underline, contents, _currentBaseline, _currentFontSize, _currentBackgroundColor, _currentFont, _currentFontFamily, underline ? _currentUnderlineStyle : OfficeIMO.Drawing.OfficeTextDecorationStyle.None, _currentStrikeStyle)); return this; }

    private PdfTextRun CreateCurrentStyleTemplate() =>
        new PdfTextRun(
            "template",
            _currentBold,
            _currentUnderline,
            _currentColor,
            _currentItalic,
            _currentStrike,
            _currentFontSize,
            _currentFont,
            baseline: _currentBaseline,
            backgroundColor: _currentBackgroundColor,
            fontFamily: _currentFontFamily,
            underlineStyle: _currentUnderlineStyle,
            strikeStyle: _currentStrikeStyle);

    private static void ValidateDecorationStyle(OfficeIMO.Drawing.OfficeTextDecorationStyle style, string parameterName) {
        if (style < OfficeIMO.Drawing.OfficeTextDecorationStyle.None || style > OfficeIMO.Drawing.OfficeTextDecorationStyle.Wavy) {
            throw new System.ArgumentOutOfRangeException(parameterName);
        }
    }

    internal RichParagraphBlock Build(PdfParagraphStyle? style = null) => new RichParagraphBlock(_runs, Align, DefaultColor, style);
}
