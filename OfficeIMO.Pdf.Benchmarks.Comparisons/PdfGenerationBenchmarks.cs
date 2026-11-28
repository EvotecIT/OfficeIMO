using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Compares complete in-memory generation of the same structured report contract.
/// Input model creation and correctness validation stay outside measured operations.
/// </summary>
// No [MemoryDiagnoser]: QuestPDF allocate in native memory, which BenchmarkDotNet does not measure.
[RankColumn]
public class PdfGenerationBenchmarks {
    private PdfBenchmarkScenario _scenario = null!;

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [GlobalSetup]
    public void Setup() {
        BenchmarkAffinityGuard.Validate();
        _scenario = PdfBenchmarkScenario.Get(Scale);
        Validate(nameof(OfficeIMO), OfficeIMO());
        Validate(nameof(QuestPDF), QuestPDF());
        Validate(nameof(MigraDoc), MigraDoc());
        Validate(nameof(IText), IText());
#if NET10_0_OR_GREATER
        Validate(nameof(OnePdfLayout), OnePdfLayout());
#endif
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public byte[] OfficeIMO() => OfficeImoPdfGenerator.Generate(_scenario);

    [Benchmark]
    [BenchmarkCategory("library:QuestPDF", "case:Default")]
    public byte[] QuestPDF() => QuestPdfGenerator.Generate(_scenario);

    [Benchmark]
    [BenchmarkCategory("library:MigraDoc", "case:Default")]
    public byte[] MigraDoc() => MigraDocPdfGenerator.Generate(_scenario);

    [Benchmark]
    [BenchmarkCategory("library:iText", "case:Default")]
    public byte[] IText() => ITextPdfGenerator.Generate(_scenario);

#if NET10_0_OR_GREATER
    /// <summary>OnePDF.Layout writes a tagged PDF by default; the other lanes write untagged or lightly tagged output.</summary>
    [Benchmark]
    [BenchmarkCategory("library:OnePDF", "case:Layout")]
    public byte[] OnePdfLayout() => OnePdfComparisonEngines.Generate(_scenario);

#endif
    // Reopens the PDF with PdfPig: page count, report markers and every page's title, narrative and table rows.
    private void Validate(string engine, byte[] bytes) =>
        PdfBenchmarkValidation.Record(
            PdfBenchmarkValidation.LibraryId(engine),
            PdfBenchmarkValidation.CaseId(engine),
            $"Scale={Scale}; reopened with PdfPig, {_scenario.PageCount} pages with every title, narrative and table row",
            () => PdfBenchmarkValidation.ValidateGenerated(bytes, _scenario, engine));
}
