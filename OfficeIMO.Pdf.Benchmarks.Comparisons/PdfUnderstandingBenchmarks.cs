using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Measures the canonical Fast semantic-read route without document-wide enrichment.</summary>
[MemoryDiagnoser]
public class PdfStructuredReadFastBenchmarks {
    private readonly PdfStructuredReadBenchmarkScenario _scenario = new();

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [GlobalSetup]
    public void Setup() => _scenario.Setup(Scale, PdfReadProfile.Fast);

    [Benchmark]
    public PdfStructuredReadObservation FastLoadAndRead() => _scenario.Read();
}

/// <summary>Measures the complete Structured semantic-read route with document-wide enrichment.</summary>
[MemoryDiagnoser]
public class PdfStructuredReadCompleteBenchmarks {
    private readonly PdfStructuredReadBenchmarkScenario _scenario = new();

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [GlobalSetup]
    public void Setup() => _scenario.Setup(Scale, PdfReadProfile.Structured);

    [Benchmark]
    public PdfStructuredReadObservation StructuredLoadAndRead() => _scenario.Read();
}

internal sealed class PdfStructuredReadBenchmarkScenario {
    private byte[] _pdf = null!;
    private PdfReadOptions _options = null!;

    internal void Setup(PdfBenchmarkScale scale, PdfReadProfile profile) {
        PdfUnderstandingBenchmarkCorpus corpus = PdfUnderstandingBenchmarkCorpusFactory.Create(scale);
        _pdf = corpus.Pdf;
        _options = PdfUnderstandingBenchmarkReadOptions.Create(profile, corpus.Pages.Count);

        PdfDocumentReadResult result = ReadDocument();
        PdfStructuredReadObservation observation = PdfUnderstandingBenchmarkValidation.Observe(result);
        if (observation.PageCount != corpus.Pages.Count || observation.TableCount < corpus.Pages.Count) {
            throw new InvalidDataException(
                $"Canonical {profile} read produced {observation.PageCount}/{observation.TableCount} pages/tables for {corpus.Pages.Count} expected pages.");
        }

        if (profile == PdfReadProfile.Structured) {
            PdfSemanticCorrectnessObservation correctness = PdfUnderstandingBenchmarkValidation.Evaluate(result, corpus);
            PdfUnderstandingBenchmarkValidation.RequireDeterministicQuality(correctness);
        }
    }

    internal PdfStructuredReadObservation Read() =>
        PdfUnderstandingBenchmarkValidation.Observe(ReadDocument());

    private PdfDocumentReadResult ReadDocument() {
        PdfDocument document = PdfDocument.Load(_pdf);
        return document.Read(_options);
    }
}

internal static class PdfUnderstandingBenchmarkReadOptions {
    private const long DefaultDocumentWorkUnits = 10_000_000L;

    internal static PdfReadOptions Create(PdfReadProfile profile, int pageCount) {
        // The labelled 100-page fixture intentionally exercises quadratic document-wide
        // similarity work. Raise only this benchmark-owned ceiling; production defaults remain bounded.
        long documentWorkUnits = Math.Max(
            DefaultDocumentWorkUnits,
            checked((long)pageCount * pageCount * 20_000L));
        return new PdfReadOptions {
            Profile = profile,
            Pipeline = new PdfUnderstandingPipelineOptions {
                MaxDocumentWorkUnits = documentWorkUnits
            }
        };
    }
}
