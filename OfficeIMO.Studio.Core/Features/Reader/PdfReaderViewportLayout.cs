namespace OfficeIMO.Studio.Features.Reader;

/// <summary>Shared page-spread and grid rules for document and independent-pane presentations.</summary>
internal static class PdfReaderViewportLayout {
    internal static int GridColumnCount(double availableWidth) => availableWidth >= 980 ? 4 : availableWidth >= 680 ? 3 : 2;

    internal static (int StartIndex, int Count) Spread(int pageNumber, int pageCount) {
        if (pageCount <= 0) return (0, 0);
        int selected = Math.Clamp(pageNumber, 1, pageCount);
        if (selected == 1) return (0, 1);
        int start = selected % 2 == 0 ? selected - 1 : selected - 2;
        return (start, Math.Min(2, pageCount - start));
    }
}
