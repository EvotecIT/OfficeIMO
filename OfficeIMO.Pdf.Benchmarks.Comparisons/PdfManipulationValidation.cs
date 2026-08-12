using UglyToad.PdfPig.DocumentLayoutAnalysis.TextExtractor;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

internal static class PdfManipulationValidation {
    internal static void Validate(
        IReadOnlyList<byte[]> outputs,
        IReadOnlyList<IReadOnlyList<PdfExpectedPage>> expectedPages,
        string engine) {
        if (outputs.Count != expectedPages.Count) {
            throw new InvalidDataException($"{engine} produced {outputs.Count} documents; expected {expectedPages.Count}.");
        }

        for (int outputIndex = 0; outputIndex < outputs.Count; outputIndex++) {
            byte[] bytes = outputs[outputIndex];
            if (bytes.Length < 5 || !bytes.AsSpan(0, 4).SequenceEqual("%PDF"u8)) {
                throw new InvalidDataException($"{engine} output {outputIndex + 1} does not have a PDF header.");
            }

            using UglyToad.PdfPig.PdfDocument document = UglyToad.PdfPig.PdfDocument.Open(bytes);
            IReadOnlyList<PdfExpectedPage> expectedOutputPages = expectedPages[outputIndex];
            if (document.NumberOfPages != expectedOutputPages.Count) {
                throw new InvalidDataException(
                    $"{engine} output {outputIndex + 1} has {document.NumberOfPages} pages; expected {expectedOutputPages.Count}.");
            }

            int pageIndex = 0;
            foreach (var page in document.GetPages()) {
                string actual = ContentOrderTextExtractor.GetText(page);
                PdfBenchmarkValidation.ValidatePageContent(
                    actual,
                    expectedOutputPages[pageIndex],
                    $"{engine} output {outputIndex + 1}, page {pageIndex + 1}");

                pageIndex++;
            }
        }
    }
}
