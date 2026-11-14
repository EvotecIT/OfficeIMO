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
    public byte[] OfficeIMO() => OfficeImoPdfGenerator.Generate(_scenario);

    [Benchmark]
    public byte[] QuestPDF() => QuestPdfGenerator.Generate(_scenario);

    [Benchmark]
    public byte[] MigraDoc() => MigraDocPdfGenerator.Generate(_scenario);

    [Benchmark]
    public byte[] IText() => ITextPdfGenerator.Generate(_scenario);

#if NET10_0_OR_GREATER
    /// <summary>OnePDF.Layout writes a tagged PDF by default; the other lanes write untagged or lightly tagged output.</summary>
    [Benchmark]
    public byte[] OnePdfLayout() => OnePdfComparisonEngines.Generate(_scenario);

#endif
    private void Validate(string engine, byte[] bytes) =>
        PdfBenchmarkValidation.ValidateGenerated(bytes, _scenario, engine);
}
