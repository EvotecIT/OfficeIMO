namespace OfficeIMO.Pdf;

internal sealed class PageBreakBlock : IPdfBlock {
    internal PageBreakBlock(bool preserveEmptyPage = false) => PreserveEmptyPage = preserveEmptyPage;
    internal bool PreserveEmptyPage { get; }
}

