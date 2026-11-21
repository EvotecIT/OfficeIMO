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
            int outputPageCount = lastPage - firstPage + 1;
            outputs.Add(CopyITextPages(
                input,
                Enumerable.Range(firstPage, outputPageCount),
                outputPageCount));
        }

        return outputs.ToArray();
    }

    internal static byte[] MergeWithIText(byte[][] sources) {
        using var outputStream = new MemoryStream();
        int outputPageCount = 0;
        using (var writer = new iText.Kernel.Pdf.PdfWriter(outputStream))
        using (var output = new iText.Kernel.Pdf.PdfDocument(writer)) {
            foreach (byte[] source in sources) {
                using var sourceStream = new MemoryStream(source, writable: false);
                using var reader = new iText.Kernel.Pdf.PdfReader(sourceStream);
                using var input = new iText.Kernel.Pdf.PdfDocument(reader);
                int inputPageCount = input.GetNumberOfPages();
                input.CopyPagesTo(1, inputPageCount, output);
                outputPageCount += inputPageCount;
            }
        }

        byte[] bytes = outputStream.ToArray();
        ValidateITextReadback(bytes, outputPageCount);
        return bytes;
    }

    internal static byte[] SelectWithIText(byte[] source, int[] pageNumbers) {
        using var sourceStream = new MemoryStream(source, writable: false);
        using var reader = new iText.Kernel.Pdf.PdfReader(sourceStream);
        using var input = new iText.Kernel.Pdf.PdfDocument(reader);
        return CopyITextPages(input, pageNumbers, pageNumbers.Length);
    }

    private static byte[] CopyITextPages(
        iText.Kernel.Pdf.PdfDocument input,
        IEnumerable<int> pageNumbers,
        int outputPageCount) {
        using var outputStream = new MemoryStream();
        using (var writer = new iText.Kernel.Pdf.PdfWriter(outputStream))
        using (var output = new iText.Kernel.Pdf.PdfDocument(writer)) {
            foreach (int pageNumber in pageNumbers) {
                input.CopyPagesTo(pageNumber, pageNumber, output);
            }
        }

        byte[] bytes = outputStream.ToArray();
        ValidateITextReadback(bytes, outputPageCount);
        return bytes;
    }

    private static void ValidateITextReadback(byte[] bytes, int outputPageCount) {
        using var stream = new MemoryStream(bytes, writable: false);
        using var reader = new iText.Kernel.Pdf.PdfReader(stream);
        using var document = new iText.Kernel.Pdf.PdfDocument(reader);
        int actualPageCount = document.GetNumberOfPages();
        if (actualPageCount != outputPageCount) {
            throw new InvalidDataException(
                $"iText post-save validation found {actualPageCount} pages; expected {outputPageCount}.");
        }
    }
}
