namespace OfficeIMO.Pdf;

/// <summary>A segment of header or footer content, either literal text, styled text, or a page token.</summary>
public sealed class PdfFooterSegment {
    /// <summary>Segment kind (text or token).</summary>
    public PdfFooterSegmentKind Kind { get; }
    /// <summary>Literal text used when <see cref="Kind"/> is <see cref="PdfFooterSegmentKind.Text"/>.</summary>
    public string? Text { get; }
    /// <summary>Optional visual styling for this segment.</summary>
    /// <remarks>Header/footer rich text supports text styling only. Links, inline visuals, and paragraph tabs are not accepted.</remarks>
    public PdfTextRun? StyledRun { get; }

    /// <summary>Creates a new footer segment.</summary>
    public PdfFooterSegment(PdfFooterSegmentKind kind, string? text = null) {
        if (kind != PdfFooterSegmentKind.Text &&
            kind != PdfFooterSegmentKind.PageNumber &&
            kind != PdfFooterSegmentKind.TotalPages &&
            kind != PdfFooterSegmentKind.DocumentPages) {
            throw new System.ArgumentOutOfRangeException(nameof(kind), "Footer segments must use a supported segment kind.");
        }

        if (kind == PdfFooterSegmentKind.Text && text == null) {
            throw new System.ArgumentNullException(nameof(text), "Footer text segments cannot be null.");
        }

        Kind = kind;
        Text = text;
    }

    private PdfFooterSegment(PdfFooterSegmentKind kind, string? text, PdfTextRun styledRun) {
        ValidateStyledRun(styledRun, nameof(styledRun));
        Kind = kind;
        Text = text;
        StyledRun = styledRun;
    }

    /// <summary>Creates a styled literal text segment.</summary>
    /// <param name="run">Visual text run to render.</param>
    public static PdfFooterSegment RichText(PdfTextRun run) {
        Guard.NotNull(run, nameof(run));
        return new PdfFooterSegment(PdfFooterSegmentKind.Text, run.Text, run);
    }

    /// <summary>Creates a current-page token using the supplied visual text style.</summary>
    /// <param name="style">Text run whose visual styling is applied; its text is ignored.</param>
    public static PdfFooterSegment PageNumber(PdfTextRun style) {
        Guard.NotNull(style, nameof(style));
        return new PdfFooterSegment(PdfFooterSegmentKind.PageNumber, null, style);
    }

    /// <summary>Creates a total-pages token using the supplied visual text style.</summary>
    /// <param name="style">Text run whose visual styling is applied; its text is ignored.</param>
    public static PdfFooterSegment TotalPages(PdfTextRun style) {
        Guard.NotNull(style, nameof(style));
        return new PdfFooterSegment(PdfFooterSegmentKind.TotalPages, null, style);
    }

    /// <summary>Creates a document-wide page-count token using the supplied visual text style.</summary>
    /// <param name="style">Text run whose visual styling is applied; its text is ignored.</param>
    public static PdfFooterSegment DocumentPages(PdfTextRun style) {
        Guard.NotNull(style, nameof(style));
        return new PdfFooterSegment(PdfFooterSegmentKind.DocumentPages, null, style);
    }

    internal static void ValidateStyledRun(PdfTextRun run, string paramName) {
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

