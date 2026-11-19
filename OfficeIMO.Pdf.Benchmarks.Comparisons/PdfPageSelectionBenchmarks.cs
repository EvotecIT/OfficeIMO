using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Extracts a deterministic non-contiguous page selection in reverse order.</summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfPageSelectionBenchmarks {
    private byte[] _source = null!;
    private int[] _pageNumbers = null!;
    private IReadOnlyList<IReadOnlyList<PdfExpectedPage>> _expectedPages = null!;

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High, PdfBenchmarkScale.VeryHigh)]
    public PdfBenchmarkScale Scale { get; set; }

    [Params(PdfBenchmarkProducer.OfficeIMO, PdfBenchmarkProducer.IText)]
    public PdfBenchmarkProducer Producer { get; set; }

    [Params(PdfManipulationDensity.Standard, PdfManipulationDensity.Dense)]
    public PdfManipulationDensity Density { get; set; }

    [GlobalSetup]
    public void Setup() {
        PdfManipulationScenario scenario = PdfManipulationScenario.Get(Scale);
        PdfBenchmarkScenario sourceScenario = scenario.SourceDocument(Density);
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
#if NET10_0_OR_GREATER
        Validate(nameof(OnePdf), OnePdf());
#endif
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public byte[] OfficeIMO() => PdfManipulationEngines.SelectWithOfficeImo(_source, _pageNumbers);

    [Benchmark]
    [BenchmarkCategory("library:iText", "case:Default")]
    public byte[] IText() => PdfManipulationEngines.SelectWithIText(_source, _pageNumbers);

    [Benchmark]
    [BenchmarkCategory("library:PDFsharp", "case:Default")]
    public byte[] PdfSharp() => PdfManipulationEngines.SelectWithPdfSharp(_source, _pageNumbers);

#if NET10_0_OR_GREATER
    [Benchmark]
    [BenchmarkCategory("library:OnePDF", "case:Default")]
    public byte[] OnePdf() => PdfManipulationEngines.SelectWithOnePdf(_source, _pageNumbers);
#endif

    // Reopens the selection with PdfPig: page count and each page's title, narrative and table rows in the requested
    // (reverse, non-contiguous) order.
    private void Validate(string engine, byte[] output) =>
        PdfBenchmarkValidation.Record(
            PdfBenchmarkValidation.LibraryId(engine),
            "Default",
            $"Density={Density}; Producer={Producer}; Scale={Scale}; {_pageNumbers.Length} selected pages reopened with PdfPig in requested order",
            () => PdfManipulationValidation.Validate(new[] { output }, _expectedPages, engine + " selecting pages from " + Producer));
}
