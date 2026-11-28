namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static partial class PdfManipulationEngines {
    internal static byte[][] SplitWithIText(byte[] source, int pagesPerDocument) {
        using var sourceStream = new MemoryStream(source, writable: false);
        using var reader = new iText.Kernel.Pdf.PdfReader(sourceStream);
        using var input = new iText.Kernel.Pdf.PdfDocument(reader);
        int pageCount = input.GetNumberOfPages();
        var outputs = new List<byte[]>((pageCount + pagesPerDocument - 1) / pagesPerDocument);
        for (int firstPage = 1; firstPage <= pageCount; firstPage += pagesPerDocument) {
            int lastPage = Math.Min(firstPage + pagesPerDocument - 1, pageCount);
            outputs.Add(CopyITextPages(input, Enumerable.Range(firstPage, lastPage - firstPage + 1)));
        }

        return outputs.ToArray();
    }

    internal static byte[] MergeWithIText(byte[][] sources) {
        using var outputStream = new MemoryStream();
        using (var writer = new iText.Kernel.Pdf.PdfWriter(outputStream))
        using (var output = new iText.Kernel.Pdf.PdfDocument(writer)) {
            foreach (byte[] source in sources) {
                using var sourceStream = new MemoryStream(source, writable: false);
                using var reader = new iText.Kernel.Pdf.PdfReader(sourceStream);
                using var input = new iText.Kernel.Pdf.PdfDocument(reader);
                input.CopyPagesTo(1, input.GetNumberOfPages(), output);
            }
        }

        return outputStream.ToArray();
    }

    internal static byte[] SelectWithIText(byte[] source, int[] pageNumbers) {
        using var sourceStream = new MemoryStream(source, writable: false);
        using var reader = new iText.Kernel.Pdf.PdfReader(sourceStream);
        using var input = new iText.Kernel.Pdf.PdfDocument(reader);
        return CopyITextPages(input, pageNumbers);
    }

    private static byte[] CopyITextPages(iText.Kernel.Pdf.PdfDocument input, IEnumerable<int> pageNumbers) {
        using var outputStream = new MemoryStream();
        using (var writer = new iText.Kernel.Pdf.PdfWriter(outputStream))
        using (var output = new iText.Kernel.Pdf.PdfDocument(writer)) {
            foreach (int pageNumber in pageNumbers) {
                input.CopyPagesTo(pageNumber, pageNumber, output);
            }
        }

        return outputStream.ToArray();
    }

}
