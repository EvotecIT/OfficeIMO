using BenchmarkDotNet.Attributes;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Measures public page import at each insertion boundary with validated output readback.</summary>
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

        PdfExpectedPage[] targetPages = Enumerable.Range(1, TargetPageCount)
            .Select(page => PdfBenchmarkValidation.ExpectedPage(targetScenario, page))
            .ToArray();
        PdfExpectedPage sourcePage = PdfBenchmarkValidation.ExpectedPage(sourceScenario, 1);
        IReadOnlyList<PdfExpectedPage> middlePages = new[] {
            PdfBenchmarkValidation.ExpectedPage(targetScenario, 1),
            sourcePage
        }.Concat(targetPages.Skip(1))
            .ToArray();
        PdfManipulationValidation.Validate(
            new[] { InsertAtStart() },
            new[] { (IReadOnlyList<PdfExpectedPage>)new[] { sourcePage }.Concat(targetPages).ToArray() },
            nameof(InsertAtStart));
        PdfManipulationValidation.Validate(
            new[] { InsertAfterFirstPage() },
            new[] { middlePages },
            nameof(InsertAfterFirstPage));
        PdfManipulationValidation.Validate(
            new[] { InsertAtEnd() },
            new[] { (IReadOnlyList<PdfExpectedPage>)targetPages.Concat(new[] { sourcePage }).ToArray() },
            nameof(InsertAtEnd));
    }

    [Benchmark]
    public byte[] InsertAtStart() => PdfDocument.Load(_target).Pages.Insert(1, _source).ToBytes();

    [Benchmark]
    public byte[] InsertAfterFirstPage() => PdfDocument.Load(_target).Pages.Insert(2, _source).ToBytes();

    [Benchmark]
    public byte[] InsertAtEnd() => PdfDocument.Load(_target).Pages.Insert(TargetPageCount + 1, _source).ToBytes();
}
