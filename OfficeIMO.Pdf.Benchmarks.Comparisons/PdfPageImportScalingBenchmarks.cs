using BenchmarkDotNet.Attributes;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Measures a public page import through extraction, insertion, and validated output readback.</summary>
[MemoryDiagnoser]
public class PdfPageImportScalingBenchmarks {
    private byte[] _target = null!;
    private byte[] _source = null!;

    [Params(5, 6, 50, 500)]
    public int TargetPageCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        var targetScenario = new PdfBenchmarkScenario(
            PdfBenchmarkScale.High, "Page import packet", TargetPageCount,
            RowsPerPage: 4, ParagraphsPerPage: 1, DocumentNumber: 1);
        var sourceScenario = targetScenario with { DocumentNumber = 2, PageCount = 1 };
        _target = PdfDocumentGenerators.Generate(PdfBenchmarkProducer.OfficeIMO, targetScenario);
        _source = PdfDocumentGenerators.Generate(PdfBenchmarkProducer.OfficeIMO, sourceScenario);
        PdfBenchmarkValidation.ValidateGenerated(_target, targetScenario, "page import target");
        PdfBenchmarkValidation.ValidateGenerated(_source, sourceScenario, "page import source");

        IReadOnlyList<PdfExpectedPage> expectedPages = new[] {
            PdfBenchmarkValidation.ExpectedPage(targetScenario, 1),
            PdfBenchmarkValidation.ExpectedPage(sourceScenario, 1)
        }.Concat(Enumerable.Range(2, TargetPageCount - 1)
            .Select(page => PdfBenchmarkValidation.ExpectedPage(targetScenario, page)))
            .ToArray();
        PdfManipulationValidation.Validate(
            new[] { InsertAfterFirstPage() },
            new[] { expectedPages },
            nameof(InsertAfterFirstPage));
    }

    [Benchmark]
    public byte[] InsertAfterFirstPage() => PdfDocument.Load(_target).Pages.Insert(2, _source).ToBytes();
}
