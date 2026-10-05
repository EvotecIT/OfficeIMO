namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static int ResolveNextVisiblePageNumber(int emittedPageCount, bool firstPageOfGroup,
        int previousVisiblePageNumber, PdfOptions options) =>
        emittedPageCount == 0 || (firstPageOfGroup && options.HasExplicitPageNumberStart)
            ? options.HasExplicitPageNumberStart ? options.PageNumberStart : 1
            : previousVisiblePageNumber + 1;
}
