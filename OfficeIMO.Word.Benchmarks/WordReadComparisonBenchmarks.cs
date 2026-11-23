using BenchmarkDotNet.Attributes;
using OfficeIMO.Benchmarks;
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

    [Params(100, 1000, 10000)]
    public int ItemCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        SetupOfficeAndOpenXml();
        if (BenchmarkLibrarySelection.IsEnabled("docx")) EnsureExpected("DocX", "Default", nameof(DocX), DocX());
        EnsureExpected("NPOI", "Default", nameof(NPOI), NPOI());
#if NET10_0_OR_GREATER
        EnsureExpected("OneDocument.Documents", "Default", nameof(OneDocument), OneDocument());
        EnsureExpected("OneDocument.Documents", "ExtractText", nameof(OneDocumentExtractText), OneDocumentExtractText());
#endif
    }

    internal void SetupOfficeAndOpenXml() {
        _fixture = WordBenchmarkCorpus.CreateParagraphFixture(ItemCount);
        _expected = WordBenchmarkCorpus.ObserveExpectedParagraphs(ItemCount);
        EnsureExpected("OfficeIMO", "Default", nameof(OfficeIMO), OfficeIMO());
        EnsureExpected("Open XML SDK", "Default", nameof(OpenXmlSdk), OpenXmlSdk());
        EnsureExpected("OfficeIMO", "PackageReader", nameof(OfficeIMOPackageReader), OfficeIMOPackageReader());
        EnsureExpected("OfficeIMO", "StreamedBodyText", nameof(OfficeIMOStreamedBodyText), OfficeIMOStreamedBodyText());
        EnsureExpected("OfficeIMO", "ReaderText", nameof(OfficeIMOReaderText), OfficeIMOReaderText());
        EnsureExpected("OfficeIMO", "ReaderStreamedText", nameof(OfficeIMOReaderStreamedText), OfficeIMOReaderStreamedText());
    }

    internal int InputBytes => _fixture.Length;

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public WordReadObservation OfficeIMO() {
        using var stream = new MemoryStream(_fixture, writable: false);
        using WordDocument document = WordDocument.Load(stream);
        var observation = WordReadObservation.Empty;
        foreach (WordParagraph paragraph in document.Paragraphs) {
            observation = observation.Add(paragraph.Text);
        }
        return observation;
    }

    /// <summary>OfficeIMO.Reader's text-and-metadata projection: the plain-text route for ingestion and search.</summary>
    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:ReaderText")]
    public WordReadObservation OfficeIMOReaderText() => WordReaderTextWorkload.Read(_fixture);

    /// <summary>OfficeIMO.Reader's <c>Text</c> projection, which returns only text chunks.</summary>
    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:ReaderStreamedText")]
    public WordReadObservation OfficeIMOReaderStreamedText() => WordReaderTextWorkload.ReadStreamed(_fixture);

    [Benchmark]
    [BenchmarkCategory("library:DocX", "case:Default")]
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
    [BenchmarkCategory("library:NPOI", "case:Default")]
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
    [BenchmarkCategory("library:Open XML SDK", "case:Default")]
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

    /// <summary>The SDK-free Open XML package and pull reader in OfficeIMO.Core: same paragraphs as the SDK lane.</summary>
    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:PackageReader")]
    public WordReadObservation OfficeIMOPackageReader() {
        using var stream = new MemoryStream(_fixture, writable: false);
        var observation = WordReadObservation.Empty;
        foreach (string text in global::OfficeIMO.Core.Internal.OpenXml.OfficeOpenXmlTextReader.ReadWordBodyParagraphs(stream)) {
            observation = observation.Add(text);
        }
        return observation;
    }

    /// <summary>The SDK-free body-text pass behind the Reader's streamed text projection, without the Reader's chunks.</summary>
    [Benchmark]
    [BenchmarkCategory("library:OfficeIMO", "case:StreamedBodyText")]
    public WordReadObservation OfficeIMOStreamedBodyText() {
        using var stream = new MemoryStream(_fixture, writable: false);
        List<string> lines = global::OfficeIMO.Word.WordStreamedBodyText.TryReadLines(stream, int.MaxValue, 0, null, includeNotes: true, CancellationToken.None, out _)
            ?? throw new InvalidDataException("The streamed body-text pass declined the fixture.");
        var observation = WordReadObservation.Empty;
        foreach (string line in lines) observation = observation.Add(line);
        return observation;
    }

#if NET10_0_OR_GREATER
    [Benchmark]
    [BenchmarkCategory("library:OneDocument.Documents", "case:Default")]
    public WordReadObservation OneDocument() => WordOneDocumentWorkloads.ReadParagraphs(_fixture);

    /// <summary>OneDocument's streaming extractor: same text, no document model.</summary>
    [Benchmark]
    [BenchmarkCategory("library:OneDocument.Documents", "case:ExtractText")]
    public WordReadObservation OneDocumentExtractText() => WordOneDocumentWorkloads.ExtractText(_fixture);

#endif
    /// <summary>
    /// Requires the lane to observe every paragraph of the fixture with the corpus text (count, characters and an
    /// order-sensitive checksum) and records the result under the lane's declared library and case.
    /// </summary>
    private void EnsureExpected(string library, string caseName, string implementation, WordReadObservation actual) =>
        WordBenchmarkCorpus.Check(library, caseName,
            $"ItemCount={ItemCount}; {_expected.ParagraphCount} paragraphs, {_expected.CharacterCount} characters, checksum {_expected.Checksum:X16} as written",
            () => EnsureExpected(implementation, actual));

    private void EnsureExpected(string implementation, WordReadObservation actual) {
        if (actual != _expected) {
            throw new InvalidDataException(
                implementation + " observed " + actual + "; expected " + _expected + ".");
        }
    }
}
