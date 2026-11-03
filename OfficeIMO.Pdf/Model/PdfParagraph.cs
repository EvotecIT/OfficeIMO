namespace OfficeIMO.Pdf;

/// <summary>
/// A paragraph of styled text runs in the document flow.
/// The paragraph stays editable until the document is laid out: runs added or formatted later are rendered.
/// </summary>
public sealed class PdfParagraph {
    private readonly System.Collections.Generic.List<PdfTextRun> _runs = new();
    private readonly PdfDocument? _owner;

    internal PdfParagraph(OfficeAlignment alignment, global::OfficeIMO.OfficeColor? color, PdfDocument? owner = null) {
        Alignment = Guard.ParagraphAlign(alignment, nameof(alignment), "Paragraph");
        FallbackColor = color;
        RunDefaults = new PdfTextRunDefaults(color);
        _owner = owner;
    }

    /// <summary>Paragraph alignment.</summary>
    public OfficeAlignment Alignment { get; }

    /// <summary>Formatting copied into each run that this paragraph creates afterwards.</summary>
    public PdfTextRunDefaults RunDefaults { get; }

    /// <summary>Runs in this paragraph, in reading order.</summary>
    public System.Collections.Generic.IReadOnlyList<PdfTextRun> Runs => _runs;

    /// <summary>Colour used by runs that carry no colour of their own, such as runs added with <see cref="AddRuns"/>.</summary>
    internal global::OfficeIMO.OfficeColor? FallbackColor { get; }

    /// <summary>Set when the paragraph is laid out in flow that cannot place page footnotes.</summary>
    internal bool RejectsFootnotes { get; set; }

    /// <summary>Appends a text run formatted from <see cref="RunDefaults"/> and returns it.</summary>
    public PdfTextRun AddText(string text) => AddRun(CreateRun(text));

    /// <summary>Appends an explicit line break formatted from <see cref="RunDefaults"/>.</summary>
    public PdfTextRun AddLineBreak() => AddText("\n");

    /// <summary>Appends an aligned paragraph tab formatted from <see cref="RunDefaults"/>.</summary>
    public PdfTextRun AddTab(PdfTabLeaderStyle leader = PdfTabLeaderStyle.None, PdfTabAlignment alignment = PdfTabAlignment.Left) =>
        AddRun(CreateRun("\t", leader, alignment));

    /// <summary>Appends a hyperlink run. It takes its size, font, background, baseline and decoration patterns from <see cref="RunDefaults"/>.</summary>
    /// <param name="text">Link text.</param>
    /// <param name="uri">Absolute URI or catalog-base-relative URI to open.</param>
    /// <param name="color">Optional link colour; defaults to <see cref="PdfTextRunDefaults.Color"/>.</param>
    /// <param name="underline">Whether to underline the link text (default true).</param>
    /// <param name="contents">Optional link annotation contents; defaults to the link text when omitted.</param>
    public PdfTextRun AddLink(string text, string uri, global::OfficeIMO.OfficeColor? color = null, bool underline = true, string? contents = null) {
        PdfTextRunDefaults d = RunDefaults;
        return AddRun(PdfTextRun.Link(text, uri, color ?? d.Color, underline, contents, d.Baseline, d.FontSize, d.BackgroundColor, d.StandardFont, d.FontFamily,
            underline ? d.UnderlineStyle : OfficeIMO.Drawing.OfficeTextDecorationStyle.None, d.StrikeStyle));
    }

    /// <summary>Appends a hyperlink run that points to a document bookmark.</summary>
    /// <param name="text">Link text.</param>
    /// <param name="bookmarkName">Named destination created with a bookmark in the same document.</param>
    /// <param name="color">Optional link colour; defaults to <see cref="PdfTextRunDefaults.Color"/>.</param>
    /// <param name="underline">Whether to underline the link text (default true).</param>
    /// <param name="contents">Optional link annotation contents; defaults to the link text when omitted.</param>
    public PdfTextRun AddBookmarkLink(string text, string bookmarkName, global::OfficeIMO.OfficeColor? color = null, bool underline = true, string? contents = null) {
        PdfTextRunDefaults d = RunDefaults;
        return AddRun(PdfTextRun.LinkToBookmark(text, bookmarkName, color ?? d.Color, underline, contents, d.Baseline, d.FontSize, d.BackgroundColor, d.StandardFont, d.FontFamily,
            underline ? d.UnderlineStyle : OfficeIMO.Drawing.OfficeTextDecorationStyle.None, d.StrikeStyle));
    }

    /// <summary>Appends an inline visual; flexible rules take their typography from <see cref="RunDefaults"/>.</summary>
    public PdfTextRun AddInline(PdfInlineElement element) {
        Guard.NotNull(element, nameof(element));
        return AddRun(element is PdfInlineRule
            ? PdfTextRun.Inline(element, CreateStyleTemplate())
            : PdfTextRun.Inline(element));
    }

    /// <summary>Appends an image that participates in paragraph wrapping.</summary>
    public PdfTextRun AddInlineImage(byte[] imageBytes, global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, string? altText = null, OfficeIMO.Drawing.OfficeImageFit fit = OfficeIMO.Drawing.OfficeImageFit.Contain, global::OfficeIMO.OfficeLength? baselineOffset = null) =>
        AddInline(new PdfInlineImage(imageBytes, width, height, altText, fit, baselineOffset ?? global::OfficeIMO.OfficeLength.Points(0D)));

    /// <summary>Appends a filled and/or bordered box that participates in paragraph wrapping.</summary>
    public PdfTextRun AddInlineBox(global::OfficeIMO.OfficeLength width, global::OfficeIMO.OfficeLength height, global::OfficeIMO.OfficeColor? background = null, global::OfficeIMO.OfficeColor? borderColor = null, global::OfficeIMO.OfficeLength? borderWidth = null, string? altText = null, global::OfficeIMO.OfficeLength? baselineOffset = null) =>
        AddInline(new PdfInlineBox(width, height, background, borderColor, borderWidth ?? global::OfficeIMO.OfficeLength.Points(0.5D), altText, baselineOffset ?? global::OfficeIMO.OfficeLength.Points(0D)));

    /// <summary>Appends prepared runs as they are, keeping their own formatting and font selections.</summary>
    public PdfParagraph AddRuns(System.Collections.Generic.IEnumerable<PdfTextRun> runs) {
        Guard.NotNull(runs, nameof(runs));
        foreach (PdfTextRun run in runs) {
            if (run?.FootnoteReference is { } note) ValidateFootnote(note, nameof(runs));
            _runs.Add(run!);
        }
        return this;
    }

    /// <summary>Appends text planned through an embedded-font fallback set, formatted from <see cref="RunDefaults"/>.</summary>
    public PdfParagraph AddFallbackText(PdfEmbeddedFontFallbackSet fallbackSet, string text, string source = "") {
        Guard.NotNull(fallbackSet, nameof(fallbackSet));
        Guard.NotNull(text, nameof(text));
        _runs.AddRange(fallbackSet.PlanTextRuns(text, source, CreateStyleTemplate()));
        return this;
    }

    internal RichParagraphBlock CreateBlock(PdfParagraphStyle? style = null) => new RichParagraphBlock(this, style);

    private PdfTextRun CreateRun(string text, PdfTabLeaderStyle tabLeader = PdfTabLeaderStyle.None, PdfTabAlignment tabAlignment = PdfTabAlignment.Left) {
        PdfTextRunDefaults d = RunDefaults;
        return new PdfTextRun(text, d.Bold, underline: false, d.Color, d.Italic, strike: false, d.FontSize, d.StandardFont,
            baseline: d.Baseline, tabLeader: tabLeader, tabAlignment: tabAlignment, backgroundColor: d.BackgroundColor, fontFamily: d.FontFamily,
            underlineStyle: d.UnderlineStyle, strikeStyle: d.StrikeStyle);
    }

    private PdfTextRun CreateStyleTemplate() => CreateRun("template").ApplyParagraphDefaults(RunDefaults);

    private PdfTextRun AddRun(PdfTextRun run) {
        run.ApplyParagraphDefaults(RunDefaults);
        _runs.Add(run);
        return run;
    }

    private void ValidateFootnote(PdfFootnoteDefinition note, string paramName) {
        if (_owner != null && !ReferenceEquals(note.Owner, _owner)) throw new System.ArgumentException("A footnote reference must belong to the same PDF document.", paramName);
        if (RejectsFootnotes) throw new System.NotSupportedException("Page-associated footnotes require ordinary single-column flow.");
    }
}
