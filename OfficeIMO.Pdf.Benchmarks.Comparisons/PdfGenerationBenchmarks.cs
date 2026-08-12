using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Compares complete in-memory generation of the same structured report contract.
/// Input model creation and correctness validation stay outside measured operations.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfGenerationBenchmarks {
    private PdfBenchmarkScenario _scenario = null!;

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [GlobalSetup]
    public void Setup() {
        _scenario = PdfBenchmarkScenario.Get(Scale);
        Validate(nameof(OfficeIMO), OfficeIMO());
        Validate(nameof(QuestPDF), QuestPDF());
        Validate(nameof(MigraDoc), MigraDoc());
        Validate(nameof(IText), IText());
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() => OfficeImoPdfGenerator.Generate(_scenario);

    [Benchmark]
    public byte[] QuestPDF() => QuestPdfGenerator.Generate(_scenario);

    [Benchmark]
    public byte[] MigraDoc() => MigraDocPdfGenerator.Generate(_scenario);

    [Benchmark]
    public byte[] IText() => ITextPdfGenerator.Generate(_scenario);

    private void Validate(string engine, byte[] bytes) =>
        PdfBenchmarkValidation.ValidateGenerated(bytes, _scenario, engine);
}
