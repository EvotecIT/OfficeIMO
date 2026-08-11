using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Extracts a deterministic non-contiguous page selection in reverse order.</summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfPageSelectionBenchmarks {
    private byte[] _source = null!;
    private int[] _pageNumbers = null!;
    private IReadOnlyList<IReadOnlyList<PdfExpectedPage>> _expectedPages = null!;

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [Params(PdfBenchmarkProducer.OfficeIMO, PdfBenchmarkProducer.IText)]
    public PdfBenchmarkProducer Producer { get; set; }

    [GlobalSetup]
    public void Setup() {
        PdfManipulationScenario scenario = PdfManipulationScenario.Get(Scale);
        PdfBenchmarkScenario sourceScenario = scenario.SourceDocument();
        _source = PdfDocumentGenerators.Generate(Producer, sourceScenario);
        PdfBenchmarkValidation.ValidateGenerated(_source, sourceScenario, Producer.ToString());
        _pageNumbers = scenario.SelectedPages();
        _expectedPages = new[] {
            (IReadOnlyList<PdfExpectedPage>)_pageNumbers
                .Select(page => PdfBenchmarkValidation.ExpectedPage(sourceScenario, page))
                .ToArray()
        };
        Validate(nameof(OfficeIMO), OfficeIMO());
        Validate(nameof(IText), IText());
        Validate(nameof(PdfSharp), PdfSharp());
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() => PdfManipulationEngines.SelectWithOfficeImo(_source, _pageNumbers);

    [Benchmark]
    public byte[] IText() => PdfManipulationEngines.SelectWithIText(_source, _pageNumbers);

    [Benchmark]
    public byte[] PdfSharp() => PdfManipulationEngines.SelectWithPdfSharp(_source, _pageNumbers);

    private void Validate(string engine, byte[] output) =>
        PdfManipulationValidation.Validate(new[] { output }, _expectedPages, engine + " selecting pages from " + Producer);
}
