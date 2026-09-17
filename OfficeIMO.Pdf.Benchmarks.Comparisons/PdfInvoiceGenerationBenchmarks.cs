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
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() => OfficeImoPdfInvoiceGenerator.Generate(_scenario);

    [Benchmark]
    public byte[] QuestPDF() => QuestPdfInvoiceGenerator.Generate(_scenario);

    [Benchmark]
    public byte[] IText() => ITextPdfInvoiceGenerator.Generate(_scenario);

    private void Validate(string engine, byte[] pdf) => PdfInvoiceComparisonValidation.Validate(pdf, _scenario, engine);
}
