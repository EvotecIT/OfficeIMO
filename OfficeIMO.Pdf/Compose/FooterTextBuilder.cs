namespace OfficeIMO.Pdf;

/// <summary>Builder for footer text segments and tokens.</summary>
public class FooterTextBuilder {
    private readonly System.Collections.Generic.List<FooterSegment> _segments;
    internal FooterTextBuilder(System.Collections.Generic.List<FooterSegment> segs) { _segments = segs; }
    /// <summary>Adds a literal text segment to the footer.</summary>
    /// <param name="s">Text to render.</param>
    /// <returns>The same builder for chaining.</returns>
    public FooterTextBuilder Text(string s) { _segments.Add(new FooterSegment(FooterSegmentKind.Text, s)); return this; }

    /// <summary>Adds a token that renders the current page number.</summary>
    /// <returns>The same builder for chaining.</returns>
    public FooterTextBuilder CurrentPage() { _segments.Add(new FooterSegment(FooterSegmentKind.PageNumber)); return this; }

    /// <summary>Adds a token that renders the total number of pages.</summary>
    /// <returns>The same builder for chaining.</returns>
    public FooterTextBuilder TotalPages() { _segments.Add(new FooterSegment(FooterSegmentKind.TotalPages)); return this; }
}
