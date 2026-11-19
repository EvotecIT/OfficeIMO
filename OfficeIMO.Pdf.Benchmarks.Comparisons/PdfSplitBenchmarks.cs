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

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High, PdfBenchmarkScale.VeryHigh)]
    public PdfBenchmarkScale Scale { get; set; }

    [Params(PdfBenchmarkProducer.OfficeIMO, PdfBenchmarkProducer.IText)]
    public PdfBenchmarkProducer Producer { get; set; }

    [Params(PdfManipulationDensity.Standard, PdfManipulationDensity.Dense)]
    public PdfManipulationDensity Density { get; set; }

    [Params(PdfSplitWorkflow.EveryPage, PdfSplitWorkflow.Bundles)]
    public PdfSplitWorkflow Workflow { get; set; }

    [GlobalSetup]
    public void Setup() {
        _scenario = PdfManipulationScenario.Get(Scale);
        PdfBenchmarkScenario sourceScenario = _scenario.SourceDocument(Density);
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
#if NET10_0_OR_GREATER
        Validate(nameof(OnePdf), OnePdf());
#endif
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public byte[][] OfficeIMO() => PdfManipulationEngines.SplitWithOfficeImo(_source, _pagesPerDocument);

    [Benchmark]
    [BenchmarkCategory("library:iText", "case:Default")]
    public byte[][] IText() => PdfManipulationEngines.SplitWithIText(_source, _pagesPerDocument);

    [Benchmark]
    [BenchmarkCategory("library:PDFsharp", "case:Default")]
    public byte[][] PdfSharp() => PdfManipulationEngines.SplitWithPdfSharp(_source, _pagesPerDocument);

#if NET10_0_OR_GREATER
    [Benchmark]
    [BenchmarkCategory("library:OnePDF", "case:Default")]
    public byte[][] OnePdf() => PdfManipulationEngines.SplitWithOnePdf(_source, _pagesPerDocument);
#endif

    // Reopens every part with PdfPig: part count, page count per part, and each page's title, narrative and table
    // rows in source order.
    private void Validate(string engine, byte[][] outputs) =>
        PdfBenchmarkValidation.Record(
            PdfBenchmarkValidation.LibraryId(engine),
            "Default",
            $"Density={Density}; Producer={Producer}; Scale={Scale}; Workflow={Workflow}; {outputs.Length} parts reopened with PdfPig, page counts and page content in order",
            () => PdfManipulationValidation.Validate(outputs, _expectedPages, engine + " splitting " + Producer));
}
