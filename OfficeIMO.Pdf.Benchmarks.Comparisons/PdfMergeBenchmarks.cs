using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Merges many identical-shape source PDFs into one ordered document.</summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfMergeBenchmarks {
    private byte[][] _sources = null!;
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
        IReadOnlyList<(PdfBenchmarkScenario Scenario, int[] Pages)> documents = scenario.ExpectedMergeDocuments(Density);
        _sources = documents
            .Select(item => PdfDocumentGenerators.Generate(Producer, item.Scenario))
            .ToArray();
        for (int index = 0; index < _sources.Length; index++) {
            PdfBenchmarkValidation.ValidateGenerated(_sources[index], documents[index].Scenario, Producer + " source " + (index + 1));
        }

        _expectedPages = new[] {
            (IReadOnlyList<PdfExpectedPage>)documents
                .SelectMany(item => item.Pages.Select(page => PdfBenchmarkValidation.ExpectedPage(item.Scenario, page)))
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
    public byte[] OfficeIMO() => PdfManipulationEngines.MergeWithOfficeImo(_sources);

    [Benchmark]
    [BenchmarkCategory("library:iText", "case:Default")]
    public byte[] IText() => PdfManipulationEngines.MergeWithIText(_sources);

    [Benchmark]
    [BenchmarkCategory("library:PDFsharp", "case:Default")]
    public byte[] PdfSharp() => PdfManipulationEngines.MergeWithPdfSharp(_sources);

#if NET10_0_OR_GREATER
    [Benchmark]
    [BenchmarkCategory("library:OnePDF", "case:Default")]
    public byte[] OnePdf() => PdfManipulationEngines.MergeWithOnePdf(_sources);
#endif

    // Reopens the merged document with PdfPig: total page count and each page's title, narrative and table rows in
    // source-document order.
    private void Validate(string engine, byte[] output) =>
        PdfBenchmarkValidation.Record(
            PdfBenchmarkValidation.LibraryId(engine),
            "Default",
            $"Density={Density}; Producer={Producer}; Scale={Scale}; {_sources.Length} sources merged, reopened with PdfPig, {_expectedPages[0].Count} pages in order",
            () => PdfManipulationValidation.Validate(new[] { output }, _expectedPages, engine + " merging " + Producer));
}
