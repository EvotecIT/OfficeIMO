using BenchmarkDotNet.Attributes;
using OfficeIMO.Pdf;

namespace OfficeIMO.Benchmarks.EpubConversions;

[MemoryDiagnoser]
[BenchmarkCategory("Performance")]
public class EpubToPdfBenchmarks {
    private byte[] _epub = [];

    [Params(BookScale.Small, BookScale.Large, BookScale.ImageHeavy)]
    public BookScale Scale { get; set; }

    [GlobalSetup(Target = nameof(Default))]
    public void SetupDefault() => Setup(OfficeConversionProfile.Faithful);

    [GlobalSetup(Target = nameof(TextOnly))]
    public void SetupTextOnly() => Setup(OfficeConversionProfile.TextOnly);

    private void Setup(OfficeConversionProfile profile) {
        _epub = BookCorpus.CreateEpub(Scale);
        var (chapters, imagesPerChapter) = BookCorpus.Size(Scale);
        using var output = new MemoryStream();
        ConvertPdf(profile, output);
        using PdfDocument pdf = PdfDocument.Load(output.ToArray());
        string text = pdf.ExtractText();
        if (!text.Contains(BookCorpus.ChapterLabel(1), StringComparison.Ordinal) ||
            !text.Contains(BookCorpus.ChapterLabel(chapters), StringComparison.Ordinal))
            throw new InvalidDataException("The PDF did not retain the first and last chapters.");
        bool styled = profile == OfficeConversionProfile.Faithful &&
            OfficeConversionCapabilityCatalog.All.Single(route => route.Id == "epub-pdf").Fidelity == OfficeConversionFidelityKind.FixedLayout;
        if (styled && (pdf.Images.Placements().Count != chapters * imagesPerChapter || pdf.Outlines.Count != chapters))
            throw new InvalidDataException("The PDF did not retain the expected images and outline.");
    }

    [Benchmark]
    public long Default() => ConvertPdf(OfficeConversionProfile.Faithful, Stream.Null);

    [Benchmark(Baseline = true)]
    public long TextOnly() => ConvertPdf(OfficeConversionProfile.TextOnly, Stream.Null);

    private long ConvertPdf(OfficeConversionProfile profile, Stream output) {
        using var input = new MemoryStream(_epub, writable: false);
        return OfficeConverter.Convert(input, OfficeDocumentFormat.Epub, output, OfficeDocumentFormat.Pdf,
            new OfficeConversionOptions { OutputProfile = profile }).ItemsWritten;
    }
}
