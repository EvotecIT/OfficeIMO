using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Compares complete in-memory HTML parsing, paged layout, and PDF serialization
/// from one identical HTML document.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfHtmlBenchmarks {
    private PdfBenchmarkScenario _scenario = null!;
    private string _html = null!;

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [GlobalSetup]
    public void Setup() {
        _scenario = PdfBenchmarkScenario.Get(Scale);
        _html = PdfHtmlScenarioBuilder.Create(_scenario);
        Validate(nameof(OfficeIMO), OfficeIMO());
        Validate(nameof(PeachPDF), PeachPDF());
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() => OfficeImoPdfGenerator.GenerateHtml(_html);

    [Benchmark]
    public byte[] PeachPDF() => PeachPdfGenerator.Generate(_html);

    private void Validate(string engine, byte[] bytes) =>
        PdfBenchmarkValidation.ValidateGenerated(bytes, _scenario, engine);
}
