namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private static void ValidatePageLabelRanges(IReadOnlyList<PdfPageLabelRange> ranges, int pageCount) {
        var seenStartPages = new HashSet<int>();
        for (int i = 0; i < ranges.Count; i++) {
            PdfPageLabelRange range = ranges[i];
            if (range.StartPageNumber > pageCount) {
                throw new InvalidOperationException("PDF page-label range start page cannot exceed the generated page count.");
            }

            if (!seenStartPages.Add(range.StartPageNumber)) {
                throw new InvalidOperationException("PDF page-label ranges cannot contain duplicate start pages.");
            }
        }
    }
}
