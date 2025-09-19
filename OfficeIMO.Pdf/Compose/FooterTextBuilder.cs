namespace OfficeIMO.Pdf;

/// <summary>Builder for footer text segments and tokens.</summary>
public class FooterTextBuilder {
    private readonly System.Collections.Generic.List<FooterSegment> _segments;
    internal FooterTextBuilder(System.Collections.Generic.List<FooterSegment> segs) { _segments = segs; }
    public FooterTextBuilder Text(string s) { _segments.Add(new FooterSegment(FooterSegmentKind.Text, s)); return this; }
    public FooterTextBuilder CurrentPage() { _segments.Add(new FooterSegment(FooterSegmentKind.PageNumber)); return this; }
    public FooterTextBuilder TotalPages() { _segments.Add(new FooterSegment(FooterSegmentKind.TotalPages)); return this; }
}

