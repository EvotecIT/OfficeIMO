namespace OfficeIMO.Word.Benchmarks;

internal static class WordLibraryBenchmarkValidation {
    internal static void RunAll() {
        foreach (int itemCount in new[] { 100, 1000 }) {
            new WordCreateParagraphComparisonBenchmarks { ItemCount = itemCount }.Validate();
            new WordCreateReportComparisonBenchmarks { RowCount = itemCount }.Validate();
            new WordReadComparisonBenchmarks { ItemCount = itemCount }.Setup();
            new WordReplaceComparisonBenchmarks { ItemCount = itemCount }.Setup();
            new WordCreateImageComparisonBenchmarks { ImageCount = itemCount / 10 }.Validate();
            new WordTemplateFillComparisonBenchmarks { ItemCount = itemCount }.Setup();
            new WordReadTableComparisonBenchmarks { RowCount = itemCount }.Setup();
        }
        new WordCreateParagraphComparisonBenchmarks { ItemCount = 10000 }.Validate();
        new WordReadComparisonBenchmarks { ItemCount = 10000 }.Setup();
    }
}
