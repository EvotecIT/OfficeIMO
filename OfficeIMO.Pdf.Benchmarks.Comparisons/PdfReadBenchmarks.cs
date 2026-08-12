using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Compares opening a PDF, enumerating every page, and extracting the complete
/// text payload from identical bytes produced by several independent engines.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfReadBenchmarks {
    private byte[] _pdf = null!;
    private PdfBenchmarkScenario _scenario = null!;

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [Params(
        PdfBenchmarkProducer.OfficeIMO,
        PdfBenchmarkProducer.QuestPDF,
        PdfBenchmarkProducer.PeachPDF,
        PdfBenchmarkProducer.MigraDoc,
        PdfBenchmarkProducer.IText)]
    public PdfBenchmarkProducer Producer { get; set; }

    [GlobalSetup]
    public void Setup() {
        _scenario = PdfBenchmarkScenario.Get(Scale);
        _pdf = PdfDocumentGenerators.Generate(Producer, _scenario);
        PdfBenchmarkValidation.ValidateGenerated(_pdf, _scenario, Producer.ToString());
        Validate(nameof(OfficeIMO), OfficeIMO());
        Validate(nameof(PdfPig), PdfPig());
        Validate(nameof(IText), IText());
    }

    [Benchmark(Baseline = true)]
    public PdfReadObservation OfficeIMO() => PdfDocumentReaders.Read(PdfReaderEngine.OfficeIMO, _pdf);

    [Benchmark]
    public PdfReadObservation PdfPig() => PdfDocumentReaders.Read(PdfReaderEngine.PdfPig, _pdf);

    [Benchmark]
    public PdfReadObservation IText() => PdfDocumentReaders.Read(PdfReaderEngine.IText, _pdf);

    private void Validate(string engine, PdfReadObservation observation) =>
        PdfBenchmarkValidation.ValidateRead(observation, _scenario, engine + " reading " + Producer);
}
