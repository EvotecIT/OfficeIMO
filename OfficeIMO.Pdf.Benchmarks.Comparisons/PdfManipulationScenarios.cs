namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

public enum PdfSplitWorkflow {
    EveryPage,
    Bundles
}

internal sealed record PdfManipulationScenario(
    PdfBenchmarkScale Scale,
    int SourcePageCount,
    int PagesPerBundle,
    int MergeDocumentCount,
    int MergePagesPerDocument,
    int SelectedPageCount) {
    internal static PdfManipulationScenario Get(PdfBenchmarkScale scale) => scale switch {
        PdfBenchmarkScale.Easy => new(scale, SourcePageCount: 5, PagesPerBundle: 2, MergeDocumentCount: 3, MergePagesPerDocument: 2, SelectedPageCount: 3),
        PdfBenchmarkScale.Medium => new(scale, SourcePageCount: 20, PagesPerBundle: 5, MergeDocumentCount: 10, MergePagesPerDocument: 4, SelectedPageCount: 10),
        PdfBenchmarkScale.High => new(scale, SourcePageCount: 100, PagesPerBundle: 10, MergeDocumentCount: 25, MergePagesPerDocument: 4, SelectedPageCount: 25),
        _ => throw new ArgumentOutOfRangeException(nameof(scale))
    };

    internal PdfBenchmarkScenario SourceDocument(int documentNumber = 0, int? pageCount = null) =>
        new(
            Scale,
            "PDF manipulation packet",
            pageCount ?? SourcePageCount,
            RowsPerPage: 4,
            ParagraphsPerPage: 1,
            DocumentNumber: documentNumber);

    internal int[] SelectedPages() {
        var pages = new int[SelectedPageCount];
        if (SelectedPageCount == 1) {
            pages[0] = SourcePageCount;
            return pages;
        }

        for (int index = 0; index < pages.Length; index++) {
            double position = index * (SourcePageCount - 1D) / (SelectedPageCount - 1D);
            pages[index] = SourcePageCount - (int)Math.Round(position);
        }

        return pages;
    }

    internal IReadOnlyList<int[]> ExpectedSplitPages(PdfSplitWorkflow workflow) {
        int pagesPerOutput = workflow == PdfSplitWorkflow.EveryPage ? 1 : PagesPerBundle;
        var outputs = new List<int[]>();
        for (int firstPage = 1; firstPage <= SourcePageCount; firstPage += pagesPerOutput) {
            int count = Math.Min(pagesPerOutput, SourcePageCount - firstPage + 1);
            outputs.Add(Enumerable.Range(firstPage, count).ToArray());
        }

        return outputs;
    }

    internal IReadOnlyList<(PdfBenchmarkScenario Scenario, int[] Pages)> ExpectedMergeDocuments() {
        var sources = new List<(PdfBenchmarkScenario, int[])>(MergeDocumentCount);
        for (int document = 1; document <= MergeDocumentCount; document++) {
            PdfBenchmarkScenario scenario = SourceDocument(document, MergePagesPerDocument);
            sources.Add((scenario, Enumerable.Range(1, MergePagesPerDocument).ToArray()));
        }

        return sources;
    }
}
