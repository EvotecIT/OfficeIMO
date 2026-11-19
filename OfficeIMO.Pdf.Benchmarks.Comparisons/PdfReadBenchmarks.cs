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
#if !NETFRAMEWORK
        PdfBenchmarkProducer.PeachPDF,
#endif
        PdfBenchmarkProducer.MigraDoc,
        PdfBenchmarkProducer.IText)]
    public PdfBenchmarkProducer Producer { get; set; }

    [GlobalSetup]
    public void Setup() {
        _scenario = PdfBenchmarkScenario.Get(Scale);
        _pdf = PdfPreparedReadInputs.Resolve(Producer, _scenario);
        PdfBenchmarkValidation.ValidateGenerated(_pdf, _scenario, Producer.ToString());
        Validate(nameof(OfficeIMO), OfficeIMO());
        Validate(nameof(PdfPig), PdfPig());
        Validate(nameof(IText), IText());
#if NET10_0_OR_GREATER
        Validate(nameof(OnePdf), OnePdf());
        Validate(nameof(OnePdfParallel), OnePdfParallel());
#endif
    }

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public PdfReadObservation OfficeIMO() => PdfDocumentReaders.Read(PdfReaderEngine.OfficeIMO, _pdf);

    [Benchmark]
    [BenchmarkCategory("library:PdfPig", "case:Default")]
    public PdfReadObservation PdfPig() => PdfDocumentReaders.Read(PdfReaderEngine.PdfPig, _pdf);

    [Benchmark]
    [BenchmarkCategory("library:iText", "case:Default")]
    public PdfReadObservation IText() => PdfDocumentReaders.Read(PdfReaderEngine.IText, _pdf);

#if NET10_0_OR_GREATER
    [Benchmark]
    [BenchmarkCategory("library:OnePDF", "case:Default")]
    public PdfReadObservation OnePdf() => OnePdfComparisonEngines.Read(_pdf);

    /// <summary>OnePDF's default whole-document extraction, which reads up to four pages at once.</summary>
    [Benchmark]
    [BenchmarkCategory("library:OnePDF", "case:Parallel")]
    public PdfReadObservation OnePdfParallel() => OnePdfComparisonEngines.ReadParallel(_pdf);

#endif
    // Extracted text, normalised to upper-case letters and digits, must hold the page count, one report marker per
    // page and every title, narrative and table row of the source scenario as often as the source has it.
    private void Validate(string engine, PdfReadObservation observation) =>
        PdfBenchmarkValidation.Record(
            PdfBenchmarkValidation.LibraryId(engine),
            PdfBenchmarkValidation.CaseId(engine),
            $"Producer={Producer}; Scale={Scale}; {observation.PageCount} pages, every source title, narrative and table row extracted",
            () => PdfBenchmarkValidation.ValidateRead(observation, _scenario, engine + " reading " + Producer));
}
