using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Compares direct PDF composition of the same two-page branded invoice contract.</summary>
// No [MemoryDiagnoser]: QuestPDF allocate in native memory, which BenchmarkDotNet does not measure.
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
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public byte[] OfficeIMO() => OfficeImoPdfInvoiceGenerator.Generate(_scenario);

    [Benchmark]
    [BenchmarkCategory("library:QuestPDF", "case:Default")]
    public byte[] QuestPDF() => QuestPdfInvoiceGenerator.Generate(_scenario);

    [Benchmark]
    [BenchmarkCategory("library:iText", "case:Default")]
    public byte[] IText() => ITextPdfInvoiceGenerator.Generate(_scenario);

#if NET10_0_OR_GREATER
    /// <summary>OnePDF.Layout writes a tagged PDF by default.</summary>
    [Benchmark]
    [BenchmarkCategory("library:OnePDF", "case:Layout")]
    public byte[] OnePdfLayout() => OnePdfInvoiceGenerator.Generate(_scenario);
#endif

    // Reopens the invoice with PdfPig: two pages, every required invoice text and the undistorted 5:1 logo.
    private void Validate(string engine, byte[] pdf) =>
        PdfBenchmarkValidation.Record(
            PdfBenchmarkValidation.LibraryId(engine),
            PdfBenchmarkValidation.CaseId(engine),
            "reopened with PdfPig: 2 pages, every required invoice text, undistorted 5:1 logo",
            () => PdfInvoiceComparisonValidation.Validate(pdf, _scenario, engine));
}
