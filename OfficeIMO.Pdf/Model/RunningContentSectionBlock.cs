namespace OfficeIMO.Pdf;

/// <summary>Marks a logical source section without interrupting physical page flow.</summary>
internal sealed class RunningContentSectionBlock : IPdfBlock {
    internal RunningContentSectionBlock(int sectionId, PdfCreateOptions? continuationOptions = null) {
        SectionId = sectionId;
        ContinuationOptions = continuationOptions;
    }
    internal int SectionId { get; }
    /// <summary>Optional presentation for the next physical page; the current page retains its assigned stories and frame.</summary>
    internal PdfCreateOptions? ContinuationOptions { get; }
}
