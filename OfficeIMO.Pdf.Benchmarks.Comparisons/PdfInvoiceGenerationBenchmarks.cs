using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Compares direct PDF composition of the same two-page branded invoice contract.</summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfInvoiceGenerationBenchmarks {
    private InvoiceComparisonScenario _scenario = null!;

    [GlobalSetup]
    public void Setup() {
        BenchmarkAffinityGuard.Validate();
        _scenario = InvoiceComparisonScenario.Create();
        Validate(nameof(OfficeIMO), OfficeIMO());
        Validate(nameof(QuestPDF), QuestPDF());
        Validate(nameof(IText), IText());
#if NET10_0_OR_GREATER
        Validate(nameof(OnePdfLayout), OnePdfLayout());
#endif
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() => OfficeImoPdfInvoiceGenerator.Generate(_scenario);

    [Benchmark]
    public byte[] QuestPDF() => QuestPdfInvoiceGenerator.Generate(_scenario);

    [Benchmark]
    public byte[] IText() => ITextPdfInvoiceGenerator.Generate(_scenario);

#if NET10_0_OR_GREATER
    /// <summary>OnePDF.Layout writes a tagged PDF by default.</summary>
    [Benchmark]
    [BenchmarkCategory("library:onepdf")]
    public byte[] OnePdfLayout() => OnePdfInvoiceGenerator.Generate(_scenario);
#endif

    private void Validate(string engine, byte[] pdf) => PdfInvoiceComparisonValidation.Validate(pdf, _scenario, engine);
}
