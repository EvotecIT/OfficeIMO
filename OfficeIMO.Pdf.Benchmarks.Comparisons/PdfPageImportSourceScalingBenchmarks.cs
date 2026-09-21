using BenchmarkDotNet.Attributes;
using OfficeIMO.Pdf;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Measures importing every source page as the source grows, with complete output validation.</summary>
[MemoryDiagnoser]
public class PdfPageImportSourceScalingBenchmarks {
    private byte[] _target = null!;
    private byte[] _source = null!;

    [Params(5, 6, 50, 500)]
    public int SourcePageCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        var targetScenario = new PdfBenchmarkScenario(
            PdfBenchmarkScale.High, "Page import packet", 2,
            RowsPerPage: 4, ParagraphsPerPage: 1, DocumentNumber: 1);
        var sourceScenario = targetScenario with { DocumentNumber = 2, PageCount = SourcePageCount };
        _target = PdfDocumentGenerators.Generate(PdfBenchmarkProducer.OfficeIMO, targetScenario);
        _source = PdfDocumentGenerators.Generate(PdfBenchmarkProducer.OfficeIMO, sourceScenario);
        PdfBenchmarkValidation.ValidateGenerated(_target, targetScenario, "page import target");
        PdfBenchmarkValidation.ValidateGenerated(_source, sourceScenario, "page import source");

        IReadOnlyList<PdfExpectedPage> expectedPages = Enumerable.Range(1, targetScenario.PageCount)
            .Select(page => PdfBenchmarkValidation.ExpectedPage(targetScenario, page))
            .Concat(Enumerable.Range(1, SourcePageCount)
                .Select(page => PdfBenchmarkValidation.ExpectedPage(sourceScenario, page)))
            .ToArray();
        PdfManipulationValidation.Validate(
            new[] { AppendAllSourcePages() },
            new[] { expectedPages },
            nameof(AppendAllSourcePages));
    }

    [Benchmark]
    public byte[] AppendAllSourcePages() => PdfDocument.Load(_target).Pages.Append(_source).ToBytes();
}
