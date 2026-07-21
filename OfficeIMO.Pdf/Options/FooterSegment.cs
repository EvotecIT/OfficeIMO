namespace OfficeIMO.Pdf;

/// <summary>A segment of header or footer content, either literal text, styled text, or a page token.</summary>
public sealed class FooterSegment {
    /// <summary>Segment kind (text or token).</summary>
    public FooterSegmentKind Kind { get; }
    /// <summary>Literal text used when <see cref="Kind"/> is <see cref="FooterSegmentKind.Text"/>.</summary>
    public string? Text { get; }
    /// <summary>Optional visual styling for this segment.</summary>
    /// <remarks>Header/footer rich text supports text styling only. Links, inline visuals, and paragraph tabs are not accepted.</remarks>
    public TextRun? StyledRun { get; }

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

    private FooterSegment(FooterSegmentKind kind, string? text, TextRun styledRun) {
        ValidateStyledRun(styledRun, nameof(styledRun));
        Kind = kind;
        Text = text;
        StyledRun = styledRun;
    }

    /// <summary>Creates a styled literal text segment.</summary>
    /// <param name="run">Visual text run to render.</param>
    public static FooterSegment RichText(TextRun run) {
        Guard.NotNull(run, nameof(run));
        return new FooterSegment(FooterSegmentKind.Text, run.Text, run);
    }

    /// <summary>Creates a current-page token using the supplied visual text style.</summary>
    /// <param name="style">Text run whose visual styling is applied; its text is ignored.</param>
    public static FooterSegment PageNumber(TextRun style) {
        Guard.NotNull(style, nameof(style));
        return new FooterSegment(FooterSegmentKind.PageNumber, null, style);
    }

    /// <summary>Creates a total-pages token using the supplied visual text style.</summary>
    /// <param name="style">Text run whose visual styling is applied; its text is ignored.</param>
    public static FooterSegment TotalPages(TextRun style) {
        Guard.NotNull(style, nameof(style));
        return new FooterSegment(FooterSegmentKind.TotalPages, null, style);
    }

    internal static void ValidateStyledRun(TextRun run, string paramName) {
        Guard.NotNull(run, paramName);
        if (run.InlineElement != null) {
            throw new System.ArgumentException("PDF header/footer rich text cannot contain inline visuals. Use the header/footer image or shape APIs instead.", paramName);
        }
        if (run.LinkUri != null || run.LinkDestinationName != null) {
            throw new System.ArgumentException("PDF header/footer rich text cannot contain interactive links.", paramName);
        }
        if (run.TabLeader != PdfTabLeaderStyle.None || run.TabAlignment != PdfTabAlignment.Left || run.Text.Contains('\t')) {
            throw new System.ArgumentException("PDF header/footer rich text cannot contain paragraph tabs.", paramName);
        }
    }
}

