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
            int outputPageCount = lastPage - firstPage + 1;
            outputs.Add(CopyITextPages(
                input,
                Enumerable.Range(firstPage, outputPageCount),
                outputPageCount));
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

            outputs.Add(SavePdfSharp(output, lastPage - firstPage));
        }

        return outputs.ToArray();
    }

    internal static byte[] MergeWithOfficeImo(byte[][] sources) =>
        OfficePdfDocument.Merge(sources.Select(static source => OfficePdfDocument.Open(source))).ToBytes();

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

    internal static byte[] MergeWithPdfSharp(byte[][] sources) {
        using var output = new PdfSharpDocument();
        foreach (byte[] source in sources) {
            using var sourceStream = new MemoryStream(source, writable: false);
            using PdfSharpDocument input = PdfSharpReader.Open(sourceStream, PdfSharpOpenMode.Import);
            for (int page = 0; page < input.PageCount; page++) {
                output.AddPage(input.Pages[page]);
            }
        }

        return SavePdfSharp(output, output.PageCount);
    }

    internal static byte[] SelectWithOfficeImo(byte[] source, int[] pageNumbers) {
        OfficePdfDocument output = OfficePdfDocument.Open(source).Pages.Extract(pageNumbers);
        int? actualPageCount = output.Pipeline.Output?.PageCount;
        if (actualPageCount != pageNumbers.Length) {
            throw new InvalidDataException(
                $"OfficeIMO post-save validation found {actualPageCount?.ToString() ?? "an unreadable output"}; expected {pageNumbers.Length} pages.");
        }

        return output.ToBytes();
    }

    internal static byte[] SelectWithIText(byte[] source, int[] pageNumbers) {
        using var sourceStream = new MemoryStream(source, writable: false);
        using var reader = new iText.Kernel.Pdf.PdfReader(sourceStream);
        using var input = new iText.Kernel.Pdf.PdfDocument(reader);
        return CopyITextPages(input, pageNumbers, pageNumbers.Length);
    }

    internal static byte[] SelectWithPdfSharp(byte[] source, int[] pageNumbers) {
        using var sourceStream = new MemoryStream(source, writable: false);
        using PdfSharpDocument input = PdfSharpReader.Open(sourceStream, PdfSharpOpenMode.Import);
        using var output = new PdfSharpDocument();
        foreach (int pageNumber in pageNumbers) {
            output.AddPage(input.Pages[pageNumber - 1]);
        }

        return SavePdfSharp(output, pageNumbers.Length);
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

    private static byte[] SavePdfSharp(PdfSharpDocument document, int outputPageCount) {
        using var output = new MemoryStream();
        document.Save(output, closeStream: false);
        byte[] bytes = output.ToArray();
        ValidatePdfSharpReadback(bytes, outputPageCount);
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

    private static void ValidatePdfSharpReadback(byte[] bytes, int outputPageCount) {
        using var stream = new MemoryStream(bytes, writable: false);
        using PdfSharpDocument document = PdfSharpReader.Open(stream, PdfSharpOpenMode.Import);
        if (document.PageCount != outputPageCount) {
            throw new InvalidDataException(
                $"PDFsharp post-save validation found {document.PageCount} pages; expected {outputPageCount}.");
        }
    }
}
