using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Measures the canonical end-to-end PDF semantic read contract, including load, parsing,
/// layout recovery, table extraction, and the selected semantic profile.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfStructuredReadBenchmarks {
    private byte[] _pdf = null!;

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [GlobalSetup]
    public void Setup() {
        PdfUnderstandingBenchmarkCorpus corpus = PdfUnderstandingBenchmarkCorpusFactory.Create(Scale);
        _pdf = corpus.Pdf;

        PdfDocumentReadResult structured = Read(PdfReadProfile.Structured);
        PdfSemanticCorrectnessObservation correctness = PdfUnderstandingBenchmarkValidation.Evaluate(structured, corpus);
        PdfUnderstandingBenchmarkValidation.RequireDeterministicQuality(correctness);

        PdfStructuredReadObservation fast = PdfUnderstandingBenchmarkValidation.Observe(Read(PdfReadProfile.Fast));
        PdfStructuredReadObservation complete = PdfUnderstandingBenchmarkValidation.Observe(structured);
        if (fast.PageCount != corpus.Pages.Count || complete.PageCount != corpus.Pages.Count ||
            fast.TableCount < corpus.Pages.Count || complete.TableCount < corpus.Pages.Count) {
            throw new InvalidDataException(
                $"Canonical read validation produced Fast={fast.PageCount}/{fast.TableCount} and Structured={complete.PageCount}/{complete.TableCount} pages/tables for {corpus.Pages.Count} expected pages.");
        }
    }

    [Benchmark(Baseline = true)]
    public PdfStructuredReadObservation FastLoadAndRead() =>
        PdfUnderstandingBenchmarkValidation.Observe(Read(PdfReadProfile.Fast));

    [Benchmark]
    public PdfStructuredReadObservation StructuredLoadAndRead() =>
        PdfUnderstandingBenchmarkValidation.Observe(Read(PdfReadProfile.Structured));

    private PdfDocumentReadResult Read(PdfReadProfile profile) {
        PdfDocument document = PdfDocument.Load(_pdf);
        return document.Read(new PdfReadOptions { Profile = profile });
    }
}
