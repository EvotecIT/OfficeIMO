using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Splits identical PDF bytes into single pages or consecutive page bundles.</summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfSplitBenchmarks {
    private byte[] _source = null!;
    private PdfManipulationScenario _scenario = null!;
    private int _pagesPerDocument;
    private IReadOnlyList<IReadOnlyList<PdfExpectedPage>> _expectedPages = null!;

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [Params(PdfBenchmarkProducer.OfficeIMO, PdfBenchmarkProducer.IText)]
    public PdfBenchmarkProducer Producer { get; set; }

    [Params(PdfSplitWorkflow.EveryPage, PdfSplitWorkflow.Bundles)]
    public PdfSplitWorkflow Workflow { get; set; }

    [GlobalSetup]
    public void Setup() {
        _scenario = PdfManipulationScenario.Get(Scale);
        PdfBenchmarkScenario sourceScenario = _scenario.SourceDocument();
        _source = PdfDocumentGenerators.Generate(Producer, sourceScenario);
        PdfBenchmarkValidation.ValidateGenerated(_source, sourceScenario, Producer.ToString());
        _pagesPerDocument = Workflow == PdfSplitWorkflow.EveryPage ? 1 : _scenario.PagesPerBundle;
        _expectedPages = _scenario.ExpectedSplitPages(Workflow)
            .Select(pages => (IReadOnlyList<PdfExpectedPage>)pages
                .Select(page => PdfBenchmarkValidation.ExpectedPage(sourceScenario, page))
                .ToArray())
            .ToArray();
        Validate(nameof(OfficeIMO), OfficeIMO());
        Validate(nameof(IText), IText());
        Validate(nameof(PdfSharp), PdfSharp());
    }

    [Benchmark(Baseline = true)]
    public byte[][] OfficeIMO() => PdfManipulationEngines.SplitWithOfficeImo(_source, _pagesPerDocument);

    [Benchmark]
    public byte[][] IText() => PdfManipulationEngines.SplitWithIText(_source, _pagesPerDocument);

    [Benchmark]
    public byte[][] PdfSharp() => PdfManipulationEngines.SplitWithPdfSharp(_source, _pagesPerDocument);

    private void Validate(string engine, byte[][] outputs) =>
        PdfManipulationValidation.Validate(outputs, _expectedPages, engine + " splitting " + Producer);
}
