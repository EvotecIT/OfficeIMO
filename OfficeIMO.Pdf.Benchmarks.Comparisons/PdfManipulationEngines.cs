using OfficePdfDocument = OfficeIMO.Pdf.PdfDocument;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static partial class PdfManipulationEngines {
    internal static byte[][] SplitWithOfficeImo(byte[] source, int pagesPerDocument) {
        using var documentStream = new MemoryStream(source, writable: false);
        using OfficePdfDocument document = OfficePdfDocument.Load(documentStream);
        int sourcePageCount = document.Pages.Count;
        IReadOnlyList<OfficePdfDocument> outputs = pagesPerDocument == 1
            ? document.Pages.Split()
            : document.Pages.Split(pagesPerDocument);
        var bytes = new byte[outputs.Count][];
        try {
            for (int index = 0; index < outputs.Count; index++) {
                bytes[index] = outputs[index].ToBytes();
                int expectedPageCount = Math.Min(
                    pagesPerDocument,
                    sourcePageCount - (index * pagesPerDocument));
                ValidateOfficeImoReadback(bytes[index], expectedPageCount);
            }
            return bytes;
        } finally {
            foreach (OfficePdfDocument output in outputs) output.Dispose();
        }
    }

    internal static byte[] MergeWithOfficeImo(byte[][] sources) {
        var documents = new List<OfficePdfDocument>(sources.Length);
        try {
            foreach (byte[] source in sources) {
                using var input = new MemoryStream(source, writable: false);
                documents.Add(OfficePdfDocument.Load(input));
            }
            using OfficePdfDocument merged = OfficePdfDocument.Merge(documents);
            byte[] bytes = merged.ToBytes();
            ValidateOfficeImoReadback(bytes, merged.Pages.Count);
            return bytes;
        } finally {
            foreach (OfficePdfDocument document in documents) document.Dispose();
        }
    }

    internal static byte[] SelectWithOfficeImo(byte[] source, int[] pageNumbers) {
        using var input = new MemoryStream(source, writable: false);
        using OfficePdfDocument document = OfficePdfDocument.Load(input);
        using OfficePdfDocument output = document.Pages.Extract(pageNumbers);
        byte[] bytes = output.ToBytes();
        ValidateOfficeImoReadback(bytes, pageNumbers.Length);
        return bytes;
    }

    private static void ValidateOfficeImoReadback(byte[] bytes, int outputPageCount) {
        using var input = new MemoryStream(bytes, writable: false);
        using OfficePdfDocument document = OfficePdfDocument.Load(input);
        int actualPageCount = document.Pages.Count;
        if (actualPageCount != outputPageCount) {
            throw new InvalidDataException(
                $"OfficeIMO post-save validation found {actualPageCount} pages; expected {outputPageCount}.");
        }
    }
}
