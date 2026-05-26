namespace OfficeIMO.Pdf;

/// <summary>A segment of footer content, either literal text or a token.</summary>
public sealed class FooterSegment {
    /// <summary>Segment kind (text or token).</summary>
    public FooterSegmentKind Kind { get; }
    /// <summary>Literal text used when <see cref="Kind"/> is <see cref="FooterSegmentKind.Text"/>.</summary>
    public string? Text { get; }
    /// <summary>Creates a new footer segment.</summary>
    public FooterSegment(FooterSegmentKind kind, string? text = null) {
        if (kind != FooterSegmentKind.Text &&
            kind != FooterSegmentKind.PageNumber &&
            kind != FooterSegmentKind.TotalPages) {
            throw new System.ArgumentOutOfRangeException(nameof(kind), "Footer segments must use a supported segment kind.");
        }

        if (kind == FooterSegmentKind.Text && text == null) {
            throw new System.ArgumentNullException(nameof(text), "Footer text segments cannot be null.");
        }

        Kind = kind;
        Text = text;
    }
}

