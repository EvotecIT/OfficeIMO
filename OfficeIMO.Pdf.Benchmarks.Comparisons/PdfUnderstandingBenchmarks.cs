using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>
/// Tracks OfficeIMO's end-to-end advanced understanding workflow.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfAdvancedUnderstandingBenchmarks {
    private byte[] _pdf = null!;
    private PdfUnderstandingPipeline _pipeline = null!;

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [GlobalSetup]
    public void Setup() {
        PdfUnderstandingBenchmarkCorpus corpus = PdfUnderstandingBenchmarkCorpusFactory.Create(Scale);
        _pdf = corpus.Pdf;
        _pipeline = new PdfUnderstandingPipeline(PdfUnderstandingPipelineOptions.Advanced());

        PdfUnderstandingResult understanding = _pipeline.Run(PdfReadDocument.Open(_pdf));
        PdfUnderstandingAccuracyObservation accuracy = PdfUnderstandingBenchmarkValidation.Evaluate(understanding, corpus.Pages);
        PdfUnderstandingBenchmarkValidation.RequireCompleteLabelCoverage(accuracy);
        PdfUnderstandingBenchmarkValidation.RequireDeterministicSemanticQuality(accuracy);
    }

    [Benchmark]
    public PdfUnderstandingPerformanceObservation AdvancedUnderstanding() =>
        PdfUnderstandingBenchmarkValidation.Observe(_pipeline.Run(PdfReadDocument.Open(_pdf)));
}

/// <summary>
/// Tracks OfficeIMO's logical structure and table extraction workflow.
/// </summary>
[MemoryDiagnoser]
[RankColumn]
public class PdfLogicalStructureBenchmarks {
    private byte[] _pdf = null!;

    [Params(PdfBenchmarkScale.Easy, PdfBenchmarkScale.Medium, PdfBenchmarkScale.High)]
    public PdfBenchmarkScale Scale { get; set; }

    [GlobalSetup]
    public void Setup() {
        PdfUnderstandingBenchmarkCorpus corpus = PdfUnderstandingBenchmarkCorpusFactory.Create(Scale);
        _pdf = corpus.Pdf;

        PdfLogicalDocument logicalDocument = PdfLogicalDocument.Load(_pdf);
        PdfLogicalStructureObservation logical = PdfUnderstandingBenchmarkValidation.Observe(logicalDocument);
        PdfBinaryClassificationScore tableDetection = PdfUnderstandingBenchmarkValidation.EvaluateTableDetection(logicalDocument, corpus.Pages);
        PdfUnderstandingBenchmarkValidation.RequireDeterministicTableQuality(tableDetection);
        if (logical.PageCount != corpus.Pages.Count || logical.TableCount < corpus.Pages.Count || logical.TableCellCount == 0) {
            throw new InvalidDataException(
                $"Logical structure validation produced {logical.PageCount} pages, {logical.TableCount} tables, and {logical.TableCellCount} table cells for {corpus.Pages.Count} expected pages.");
        }
    }

    [Benchmark]
    public PdfLogicalStructureObservation LogicalStructureAndTables() =>
        PdfUnderstandingBenchmarkValidation.Observe(PdfLogicalDocument.Load(_pdf));
}
