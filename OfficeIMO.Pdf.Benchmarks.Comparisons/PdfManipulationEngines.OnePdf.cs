#if NET10_0_OR_GREATER
using OnePdfDocument = OnePdf.PdfDocument;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

// OnePDF ships net10.0 only. Sources open read-only (the default), which the library documents as
// the mode where repeated imports from one source reuse the objects already copied.
internal static partial class PdfManipulationEngines {
    internal static byte[][] SplitWithOnePdf(byte[] source, int pagesPerDocument) {
        using OnePdfDocument input = OnePdfDocument.Open(source);
        IReadOnlyList<OnePdfDocument> parts = OnePdfDocument.Split(input, pagesPerDocument);
        try {
            var outputs = new byte[parts.Count][];
            for (int index = 0; index < parts.Count; index++) {
                outputs[index] = SaveOnePdf(parts[index]);
            }

            return outputs;
        } finally {
            foreach (OnePdfDocument part in parts) part.Dispose();
        }
    }

    internal static byte[] MergeWithOnePdf(byte[][] sources) {
        var documents = new List<OnePdfDocument>(sources.Length);
        try {
            foreach (byte[] source in sources) {
                documents.Add(OnePdfDocument.Open(source));
            }

            using OnePdfDocument merged = OnePdfDocument.Merge(documents);
            return SaveOnePdf(merged);
        } finally {
            foreach (OnePdfDocument document in documents) document.Dispose();
        }
    }

    internal static byte[] SelectWithOnePdf(byte[] source, int[] pageNumbers) {
        using OnePdfDocument input = OnePdfDocument.Open(source);
        using OnePdfDocument output = OnePdfDocument.CreateEmpty();
        var pageIndexes = new int[pageNumbers.Length];
        for (int index = 0; index < pageNumbers.Length; index++) {
            pageIndexes[index] = pageNumbers[index] - 1;
        }

        output.ImportPages(input, pageIndexes);
        return SaveOnePdf(output);
    }

    private static byte[] SaveOnePdf(OnePdfDocument document) {
        using var output = new MemoryStream();
        document.Save(output);
        return output.ToArray();
    }

}
#endif
