namespace OfficeIMO.Studio.Features.Reader;

/// <summary>
/// Groups a bounded number of pages into one virtualized reader-grid row.
/// </summary>
public sealed record ReaderGridRowViewModel(IReadOnlyList<PdfPageViewModel> Pages) {
    public bool Contains(PdfPageViewModel? page) => page is not null && Pages.Contains(page);
}
