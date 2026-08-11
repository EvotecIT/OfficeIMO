using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using NpoiDocument = NPOI.XWPF.UserModel.XWPFDocument;
using XceedDocX = Xceed.Words.NET.DocX;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>Compares full traversal of the same deterministic body-paragraph corpus.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("Word", "Comparison", "Read")]
public class WordReadComparisonBenchmarks {
    private byte[] _fixture = Array.Empty<byte>();
    private WordReadObservation _expected;

    [Params(100, 1000)]
    public int ItemCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        _fixture = WordBenchmarkCorpus.CreateParagraphFixture(ItemCount);
        _expected = WordBenchmarkCorpus.ObserveExpectedParagraphs(ItemCount);
        EnsureExpected(nameof(OfficeIMO), OfficeIMO());
        EnsureExpected(nameof(DocX), DocX());
        EnsureExpected(nameof(NPOI), NPOI());
        EnsureExpected(nameof(OpenXmlSdk), OpenXmlSdk());
    }

    [Benchmark(Baseline = true)]
    public WordReadObservation OfficeIMO() {
        using var stream = new MemoryStream(_fixture, writable: false);
        using WordDocument document = WordDocument.Load(stream);
        var observation = WordReadObservation.Empty;
        foreach (WordParagraph paragraph in document.Paragraphs) {
            observation = observation.Add(paragraph.Text);
        }
        return observation;
    }

    [Benchmark]
    public WordReadObservation DocX() {
        using var stream = new MemoryStream(_fixture, writable: false);
        using XceedDocX document = XceedDocX.Load(stream);
        var observation = WordReadObservation.Empty;
        foreach (Xceed.Document.NET.Paragraph paragraph in document.Paragraphs) {
            observation = observation.Add(paragraph.Text);
        }
        return observation;
    }

    [Benchmark]
    public WordReadObservation NPOI() {
        using var stream = new MemoryStream(_fixture, writable: false);
        using var document = new NpoiDocument(stream);
        var observation = WordReadObservation.Empty;
        foreach (NPOI.XWPF.UserModel.XWPFParagraph paragraph in document.Paragraphs) {
            observation = observation.Add(paragraph.Text);
        }
        return observation;
    }

    [Benchmark]
    public WordReadObservation OpenXmlSdk() {
        using var stream = new MemoryStream(_fixture, writable: false);
        using WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: false);
        Body body = document.MainDocumentPart?.Document?.Body
            ?? throw new InvalidDataException("The benchmark fixture has no body.");
        var observation = WordReadObservation.Empty;
        foreach (Paragraph paragraph in body.Elements<Paragraph>()) {
            observation = observation.Add(paragraph.InnerText);
        }
        return observation;
    }

    private void EnsureExpected(string implementation, WordReadObservation actual) {
        if (actual != _expected) {
            throw new InvalidDataException(
                implementation + " observed " + actual + "; expected " + _expected + ".");
        }
    }
}
