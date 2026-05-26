namespace OfficeIMO.Pdf;

internal static class PdfStampPageSelection {
    internal static int[] BuildInclusivePageRange(int firstPage, int lastPage) {
        return new PdfPageRange(firstPage, lastPage).ToPageNumbers();
    }

    internal static int[] BuildInclusivePageRange(PdfPageRange pageRange) {
        return pageRange.ToPageNumbers();
    }

    internal static int[] BuildInclusivePageRanges(params PdfPageRange[] pageRanges) {
        Guard.NotNull(pageRanges, nameof(pageRanges));
        if (pageRanges.Length == 0) {
            throw new System.ArgumentException("At least one page range must be specified.", nameof(pageRanges));
        }

        var seen = new System.Collections.Generic.HashSet<int>();
        var pages = new System.Collections.Generic.List<int>();
        for (int i = 0; i < pageRanges.Length; i++) {
            int[] pageNumbers = pageRanges[i].ToPageNumbers();
            for (int j = 0; j < pageNumbers.Length; j++) {
                if (seen.Add(pageNumbers[j])) {
                    pages.Add(pageNumbers[j]);
                }
            }
        }

        return pages.ToArray();
    }
}
