using BenchmarkDotNet.Attributes;

namespace OfficeIMO.Pdf.Benchmarks.Comparisons;

/// <summary>Compares equivalent two-source merge workflows through the public PDF APIs.</summary>
[MemoryDiagnoser]
public class PdfMergeApiScalingBenchmarks {
    private byte[] _first = null!;
    private byte[] _second = null!;

    [Params(5, 6, 50, 500)]
    public int PagesPerSource { get; set; }

    [GlobalSetup]
    public void Setup() {
        var firstScenario = new PdfBenchmarkScenario(
            PdfBenchmarkScale.High, "Merge API packet", PagesPerSource,
            RowsPerPage: 4, ParagraphsPerPage: 1, DocumentNumber: 1);
        var secondScenario = firstScenario with { DocumentNumber = 2 };
        _first = PdfDocumentGenerators.Generate(PdfBenchmarkProducer.OfficeIMO, firstScenario);
        _second = PdfDocumentGenerators.Generate(PdfBenchmarkProducer.OfficeIMO, secondScenario);
        PdfBenchmarkValidation.ValidateGenerated(_first, firstScenario, "first merge source");
        PdfBenchmarkValidation.ValidateGenerated(_second, secondScenario, "second merge source");

        IReadOnlyList<PdfExpectedPage> expectedPages = Enumerable.Range(1, PagesPerSource)
            .Select(page => PdfBenchmarkValidation.ExpectedPage(firstScenario, page))
            .Concat(Enumerable.Range(1, PagesPerSource)
                .Select(page => PdfBenchmarkValidation.ExpectedPage(secondScenario, page)))
            .ToArray();
        IReadOnlyList<IReadOnlyList<PdfExpectedPage>> expected = new[] { expectedPages };
        PdfManipulationValidation.Validate(new[] { MergeDocuments() }, expected, nameof(MergeDocuments));
        PdfManipulationValidation.Validate(new[] { MergeWithBytes() }, expected, nameof(MergeWithBytes));
        PdfManipulationValidation.Validate(new[] { MergeWithDocument() }, expected, nameof(MergeWithDocument));
    }

    [Benchmark(Baseline = true)]
    public byte[] MergeDocuments() {
        using var firstInput = new MemoryStream(_first, writable: false);
        using var secondInput = new MemoryStream(_second, writable: false);
        using PdfDocument first = PdfDocument.Load(firstInput);
        using PdfDocument second = PdfDocument.Load(secondInput);
        using PdfDocument merged = PdfDocument.Merge(new[] { first, second });
        return merged.ToBytes();
    }

    [Benchmark]
    public byte[] MergeWithBytes() {
        using var input = new MemoryStream(_first, writable: false);
        using PdfDocument first = PdfDocument.Load(input);
        using PdfDocument merged = first.MergeWith(_second);
        return merged.ToBytes();
    }

    [Benchmark]
    public byte[] MergeWithDocument() {
        using var firstInput = new MemoryStream(_first, writable: false);
        using var secondInput = new MemoryStream(_second, writable: false);
        using PdfDocument first = PdfDocument.Load(firstInput);
        using PdfDocument second = PdfDocument.Load(secondInput);
        using PdfDocument merged = first.MergeWith(second);
        return merged.ToBytes();
    }
}
