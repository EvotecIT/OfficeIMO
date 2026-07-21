namespace OfficeIMO.Pdf;

/// <summary>Builder for header text segments and page tokens.</summary>
public class HeaderTextBuilder {
    private readonly System.Collections.Generic.List<FooterSegment> _segments;
    internal HeaderTextBuilder(System.Collections.Generic.List<FooterSegment> segments) { _segments = segments; }

    /// <summary>Adds a literal text segment to the header.</summary>
    /// <param name="s">Text to render.</param>
    /// <returns>The same builder for chaining.</returns>
    public HeaderTextBuilder Text(string s) {
        Guard.NotNull(s, nameof(s));
        _segments.Add(new FooterSegment(FooterSegmentKind.Text, s));
        return this;
    }

    /// <summary>Adds a visually styled text run to the header.</summary>
    /// <param name="run">Styled text to render. Interactive links, inline visuals, and paragraph tabs are not supported.</param>
    /// <returns>The same builder for chaining.</returns>
    public HeaderTextBuilder Run(TextRun run) {
        Guard.NotNull(run, nameof(run));
        _segments.Add(FooterSegment.RichText(run));
        return this;
    }

    /// <summary>Adds a token that renders the current page number.</summary>
    /// <returns>The same builder for chaining.</returns>
    public HeaderTextBuilder CurrentPage() {
        _segments.Add(new FooterSegment(FooterSegmentKind.PageNumber));
        return this;
    }

    /// <summary>Adds a current-page token with the supplied visual text style.</summary>
    /// <param name="style">Text run whose visual styling is applied; its text is ignored.</param>
    /// <returns>The same builder for chaining.</returns>
    public HeaderTextBuilder CurrentPage(TextRun style) {
        Guard.NotNull(style, nameof(style));
        _segments.Add(FooterSegment.PageNumber(style));
        return this;
    }

    /// <summary>Adds a token that renders the total number of pages.</summary>
    /// <returns>The same builder for chaining.</returns>
    public HeaderTextBuilder TotalPages() {
        _segments.Add(new FooterSegment(FooterSegmentKind.TotalPages));
        return this;
    }

    /// <summary>Adds a total-pages token with the supplied visual text style.</summary>
    /// <param name="style">Text run whose visual styling is applied; its text is ignored.</param>
    /// <returns>The same builder for chaining.</returns>
    public HeaderTextBuilder TotalPages(TextRun style) {
        Guard.NotNull(style, nameof(style));
        _segments.Add(FooterSegment.TotalPages(style));
        return this;
    }
}
