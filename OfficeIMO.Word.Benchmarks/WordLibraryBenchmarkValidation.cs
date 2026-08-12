namespace OfficeIMO.Word.Benchmarks;

internal static class WordLibraryBenchmarkValidation {
    internal static void RunAll() {
        foreach (int itemCount in new[] { 100, 1000 }) {
            new WordCreateParagraphComparisonBenchmarks { ItemCount = itemCount }.Validate();
            new WordCreateReportComparisonBenchmarks { RowCount = itemCount }.Validate();
            new WordReadComparisonBenchmarks { ItemCount = itemCount }.Setup();
            new WordReplaceComparisonBenchmarks { ItemCount = itemCount }.Setup();
        }
    }
}
