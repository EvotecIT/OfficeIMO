using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Merges many identical-shape source PDFs into one ordered document.</summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfMergeBenchmarks {
    private byte[][] _sources = null!;
    private IReadOnlyList<IReadOnlyList<PdfExpectedPage>> _expectedPages = null!;

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [Params(PdfBenchmarkProducer.OfficeIMO, PdfBenchmarkProducer.IText)]
    public PdfBenchmarkProducer Producer { get; set; }

    [GlobalSetup]
    public void Setup() {
        PdfManipulationScenario scenario = PdfManipulationScenario.Get(Scale);
        IReadOnlyList<(PdfBenchmarkScenario Scenario, int[] Pages)> documents = scenario.ExpectedMergeDocuments();
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
    }

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() => PdfManipulationEngines.MergeWithOfficeImo(_sources);

    [Benchmark]
    public byte[] IText() => PdfManipulationEngines.MergeWithIText(_sources);

    [Benchmark]
    public byte[] PdfSharp() => PdfManipulationEngines.MergeWithPdfSharp(_sources);

    private void Validate(string engine, byte[] output) =>
        PdfManipulationValidation.Validate(new[] { output }, _expectedPages, engine + " merging " + Producer);
}
