using PdfSharpDocument = PdfSharp.Pdf.PdfDocument;
using PdfSharpReader = PdfSharp.Pdf.IO.PdfReader;
using PdfSharpOpenMode = PdfSharp.Pdf.IO.PdfDocumentOpenMode;
using OfficePdfDocument = OfficeIMO.Pdf.PdfDocument;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class PdfManipulationEngines {
    internal static byte[][] SplitWithOfficeImo(byte[] source, int pagesPerDocument) {
        OfficePdfDocument document = OfficePdfDocument.Open(source);
        IReadOnlyList<OfficePdfDocument> outputs = pagesPerDocument == 1
            ? document.Pages.Split()
            : document.Pages.Split(pagesPerDocument);
        return outputs.Select(static output => output.ToBytes()).ToArray();
    }

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

    internal static byte[][] SplitWithPdfSharp(byte[] source, int pagesPerDocument) {
        using var sourceStream = new MemoryStream(source, writable: false);
        using PdfSharpDocument input = PdfSharpReader.Open(sourceStream, PdfSharpOpenMode.Import);
        var outputs = new List<byte[]>((input.PageCount + pagesPerDocument - 1) / pagesPerDocument);
        for (int firstPage = 0; firstPage < input.PageCount; firstPage += pagesPerDocument) {
            int lastPage = Math.Min(firstPage + pagesPerDocument, input.PageCount);
            using var output = new PdfSharpDocument();
            for (int page = firstPage; page < lastPage; page++) {
                output.AddPage(input.Pages[page]);
            }

            outputs.Add(SavePdfSharp(output));
        }

        return outputs.ToArray();
    }

    internal static byte[] MergeWithOfficeImo(byte[][] sources) =>
        OfficePdfDocument.Merge(sources.Select(static source => OfficePdfDocument.Open(source))).ToBytes();

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

    internal static byte[] MergeWithPdfSharp(byte[][] sources) {
        using var output = new PdfSharpDocument();
        foreach (byte[] source in sources) {
            using var sourceStream = new MemoryStream(source, writable: false);
            using PdfSharpDocument input = PdfSharpReader.Open(sourceStream, PdfSharpOpenMode.Import);
            for (int page = 0; page < input.PageCount; page++) {
                output.AddPage(input.Pages[page]);
            }
        }

        return SavePdfSharp(output);
    }

    internal static byte[] SelectWithOfficeImo(byte[] source, int[] pageNumbers) =>
        OfficePdfDocument.Open(source).Pages.Extract(pageNumbers).ToBytes();

    internal static byte[] SelectWithIText(byte[] source, int[] pageNumbers) {
        using var sourceStream = new MemoryStream(source, writable: false);
        using var reader = new iText.Kernel.Pdf.PdfReader(sourceStream);
        using var input = new iText.Kernel.Pdf.PdfDocument(reader);
        return CopyITextPages(input, pageNumbers);
    }

    internal static byte[] SelectWithPdfSharp(byte[] source, int[] pageNumbers) {
        using var sourceStream = new MemoryStream(source, writable: false);
        using PdfSharpDocument input = PdfSharpReader.Open(sourceStream, PdfSharpOpenMode.Import);
        using var output = new PdfSharpDocument();
        foreach (int pageNumber in pageNumbers) {
            output.AddPage(input.Pages[pageNumber - 1]);
        }

        return SavePdfSharp(output);
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

    private static byte[] SavePdfSharp(PdfSharpDocument document) {
        using var output = new MemoryStream();
        document.Save(output, closeStream: false);
        return output.ToArray();
    }

}
