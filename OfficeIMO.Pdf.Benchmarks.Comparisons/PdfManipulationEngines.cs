using OfficePdfDocument = OfficeIMO.Pdf.PdfDocument;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

// Every lane ends with the serialized output bytes. Post-save checks run once in each suite's setup with PdfPig,
// outside timing: reopening the output with the producing library inside the timed operation measured readers of
// different depth (PDFsharp's import mode parses every object; others only resolve the page tree).
internal static partial class PdfManipulationEngines {
    internal static byte[][] SplitWithOfficeImo(byte[] source, int pagesPerDocument) {
        using var documentStream = new MemoryStream(source, writable: false);
        using OfficePdfDocument document = OfficePdfDocument.Load(documentStream);
        IReadOnlyList<OfficePdfDocument> outputs = pagesPerDocument == 1
            ? document.Pages.Split()
            : document.Pages.Split(pagesPerDocument);
        var bytes = new byte[outputs.Count][];
        try {
            for (int index = 0; index < outputs.Count; index++) {
                bytes[index] = outputs[index].ToBytes();
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
            return merged.ToBytes();
        } finally {
            foreach (OfficePdfDocument document in documents) document.Dispose();
        }
    }

    internal static byte[] SelectWithOfficeImo(byte[] source, int[] pageNumbers) {
        using var input = new MemoryStream(source, writable: false);
        using OfficePdfDocument document = OfficePdfDocument.Load(input);
        using OfficePdfDocument output = document.Pages.Extract(pageNumbers);
        return output.ToBytes();
    }

}
