namespace OfficeIMO.Pdf;

/// <summary>Builder for footer text segments and tokens.</summary>
public sealed class PdfFooterTextBuilder {
    private readonly System.Collections.Generic.List<PdfFooterSegment> _segments;
    internal PdfFooterTextBuilder(System.Collections.Generic.List<PdfFooterSegment> segs) { _segments = segs; }
    /// <summary>Adds a literal text segment to the footer.</summary>
    /// <param name="s">Text to render.</param>
    /// <returns>The same builder for chaining.</returns>
    public PdfFooterTextBuilder Text(string s) { Guard.NotNull(s, nameof(s)); _segments.Add(new PdfFooterSegment(PdfFooterSegmentKind.Text, s)); return this; }

    /// <summary>Adds a visually styled text run to the footer.</summary>
    /// <param name="run">Styled text to render. Interactive links, inline visuals, and paragraph tabs are not supported.</param>
    /// <returns>The same builder for chaining.</returns>
    public PdfFooterTextBuilder Run(PdfTextRun run) { Guard.NotNull(run, nameof(run)); _segments.Add(PdfFooterSegment.RichText(run)); return this; }

    /// <summary>Adds a token that renders the current page number.</summary>
    /// <returns>The same builder for chaining.</returns>
    public PdfFooterTextBuilder CurrentPage() { _segments.Add(new PdfFooterSegment(PdfFooterSegmentKind.PageNumber)); return this; }

    /// <summary>Adds a current-page token with the supplied visual text style.</summary>
    /// <param name="style">Text run whose visual styling is applied; its text is ignored.</param>
    /// <returns>The same builder for chaining.</returns>
    public PdfFooterTextBuilder CurrentPage(PdfTextRun style) { Guard.NotNull(style, nameof(style)); _segments.Add(PdfFooterSegment.PageNumber(style)); return this; }

    /// <summary>Adds a token that renders the total number of pages.</summary>
    /// <returns>The same builder for chaining.</returns>
    public PdfFooterTextBuilder TotalPages() { _segments.Add(new PdfFooterSegment(PdfFooterSegmentKind.TotalPages)); return this; }

    /// <summary>Adds a total-pages token with the supplied visual text style.</summary>
    /// <param name="style">Text run whose visual styling is applied; its text is ignored.</param>
    /// <returns>The same builder for chaining.</returns>
    public PdfFooterTextBuilder TotalPages(PdfTextRun style) { Guard.NotNull(style, nameof(style)); _segments.Add(PdfFooterSegment.TotalPages(style)); return this; }
    /// <summary>Adds the page count of the complete document, including all sections.</summary>
    /// <returns>The same builder for chaining.</returns>
    public PdfFooterTextBuilder DocumentPages() { _segments.Add(new PdfFooterSegment(PdfFooterSegmentKind.DocumentPages)); return this; }

    /// <summary>Adds a document-wide page-count token with the supplied visual text style.</summary>
    /// <param name="style">Text run whose visual styling is applied; its text is ignored.</param>
    /// <returns>The same builder for chaining.</returns>
    public PdfFooterTextBuilder DocumentPages(PdfTextRun style) { Guard.NotNull(style, nameof(style)); _segments.Add(PdfFooterSegment.DocumentPages(style)); return this; }
}
