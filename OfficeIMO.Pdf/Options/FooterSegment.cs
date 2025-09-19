namespace OfficeIMO.Pdf;

/// <summary>A segment of footer content, either literal text or a token.</summary>
public sealed class FooterSegment {
    /// <summary>Segment kind (text or token).</summary>
    public FooterSegmentKind Kind { get; }
    /// <summary>Literal text used when <see cref="Kind"/> is <see cref="FooterSegmentKind.Text"/>.</summary>
    public string? Text { get; }
    /// <summary>Creates a new footer segment.</summary>
    public FooterSegment(FooterSegmentKind kind, string? text = null) { Kind = kind; Text = text; }
}

