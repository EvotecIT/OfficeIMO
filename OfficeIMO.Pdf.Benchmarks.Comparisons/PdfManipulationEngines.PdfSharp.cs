using PdfSharpDocument = PdfSharp.Pdf.PdfDocument;
using PdfSharpReader = PdfSharp.Pdf.IO.PdfReader;
using PdfSharpOpenMode = PdfSharp.Pdf.IO.PdfDocumentOpenMode;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

// OfficeIMO.Pdf.Benchmarks.FreePdf compiles this file unchanged against FreePDF,
// which keeps the PdfSharp namespaces, so both lanes do identical work.
internal static partial class PdfManipulationEngines {
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

    internal static byte[] SelectWithPdfSharp(byte[] source, int[] pageNumbers) {
        using var sourceStream = new MemoryStream(source, writable: false);
        using PdfSharpDocument input = PdfSharpReader.Open(sourceStream, PdfSharpOpenMode.Import);
        using var output = new PdfSharpDocument();
        foreach (int pageNumber in pageNumbers) {
            output.AddPage(input.Pages[pageNumber - 1]);
        }

        return SavePdfSharp(output, pageNumbers.Length);
    }

    private static byte[] SavePdfSharp(PdfSharpDocument document, int outputPageCount) {
        using var output = new MemoryStream();
        document.Save(output, closeStream: false);
        byte[] bytes = output.ToArray();
        ValidatePdfSharpReadback(bytes, outputPageCount);
        return bytes;
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
