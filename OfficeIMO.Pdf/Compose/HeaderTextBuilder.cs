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

    /// <summary>Adds a token that renders the current page number.</summary>
    /// <returns>The same builder for chaining.</returns>
    public HeaderTextBuilder CurrentPage() {
        _segments.Add(new FooterSegment(FooterSegmentKind.PageNumber));
        return this;
    }

    /// <summary>Adds a token that renders the total number of pages.</summary>
    /// <returns>The same builder for chaining.</returns>
    public HeaderTextBuilder TotalPages() {
        _segments.Add(new FooterSegment(FooterSegmentKind.TotalPages));
        return this;
    }
}
