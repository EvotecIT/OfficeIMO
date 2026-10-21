namespace OfficeIMO.Pdf;

/// <summary>An immutable, document-owned footnote composed with the ordinary PDF flow API.</summary>
/// <remarks>Reference the definition from a styled <see cref="PdfTextRun"/>. Pagination and continuation state belong to each serialization.
/// A continued note requires the same usable page width; finish the note before changing page width or horizontal margins.</remarks>
public sealed class PdfFootnoteDefinition {
    internal PdfFootnoteDefinition(PdfDocument owner, IReadOnlyList<IPdfBlock> blocks, PdfFootnoteOptions options,
        IReadOnlyList<IPdfBlock> separator, IReadOnlyList<IPdfBlock> continuationSeparator,
        IReadOnlyList<IPdfBlock> continuationNotice) {
        Owner = owner;
        Blocks = blocks;
        Placement = options.Placement;
        Separator = separator;
        ContinuationSeparator = continuationSeparator;
        ContinuationNotice = continuationNotice;
        ReserveContinuationNoticeSpaceOnFinalPage = options.ReserveContinuationNoticeSpaceOnFinalPage;
    }

    internal PdfDocument Owner { get; }
    internal IReadOnlyList<IPdfBlock> Blocks { get; }
    internal IReadOnlyList<IPdfBlock> Separator { get; }
    internal IReadOnlyList<IPdfBlock> ContinuationSeparator { get; }
    internal IReadOnlyList<IPdfBlock> ContinuationNotice { get; }
    internal bool ReserveContinuationNoticeSpaceOnFinalPage { get; }
    /// <summary>Placement selected when this definition was composed.</summary>
    public PdfFootnotePlacement Placement { get; }
}

/// <summary>Vertical placement of page-associated footnotes within the body frame.</summary>
public enum PdfFootnotePlacement {
    /// <summary>Align the accepted footnote fragment to the bottom of the page's body frame, above any running footer.</summary>
    BottomOfPage,
    /// <summary>Place the accepted footnote fragment immediately beneath the page's body text.</summary>
    BeneathText
}

/// <summary>Composition options captured by <see cref="PdfContentBuilder.DefineFootnote"/>.</summary>
public sealed class PdfFootnoteOptions {
    /// <summary>Vertical placement. The default is <see cref="PdfFootnotePlacement.BottomOfPage"/>.</summary>
    public PdfFootnotePlacement Placement { get; set; }
    /// <summary>Optional ordinary flow rendered before the first footnote on a page. No separator is invented when omitted.</summary>
    /// <remarks>The first admitted reference selects the page's separator. It remains above all notes when their authored source order differs from admission order.</remarks>
    public Action<PdfContentBuilder>? Separator { get; set; }
    /// <summary>Optional ordinary flow rendered before a continued footnote on a later page.</summary>
    public Action<PdfContentBuilder>? ContinuationSeparator { get; set; }
    /// <summary>Optional ordinary flow rendered after the page's notes when note content continues to another page.</summary>
    /// <remarks>The first admitted note selects the page's notice, as it does the separator.
    /// The notice reserves space before note fragments are accepted and is omitted on the final note page.</remarks>
    public Action<PdfContentBuilder>? ContinuationNotice { get; set; }
    /// <summary>Reserve the measured continuation-notice flow even on a complete or final note page. The default is false.</summary>
    /// <remarks>The first admitted note selects this policy together with the page's notice.
    /// This changes available space; the notice is still painted only when note content continues to another page.</remarks>
    public bool ReserveContinuationNoticeSpaceOnFinalPage { get; set; }
}
