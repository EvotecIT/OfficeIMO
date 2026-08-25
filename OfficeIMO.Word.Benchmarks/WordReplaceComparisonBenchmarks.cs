using System.Text.RegularExpressions;
using BenchmarkDotNet.Attributes;
using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Wordprocessing;
using Xceed.Document.NET;
using NpoiDocument = NPOI.XWPF.UserModel.XWPFDocument;
using XceedDocX = Xceed.Words.NET.DocX;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>Compares load, whole-document text replacement, and save over the same DOCX package.</summary>
[MemoryDiagnoser]
[BenchmarkCategory("Word", "Comparison", "Replace")]
public class WordReplaceComparisonBenchmarks {
    private byte[] _fixture = Array.Empty<byte>();

    [Params(100, 1000)]
    public int ItemCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        SetupOfficeAndOpenXml();
        WordBenchmarkCorpus.ValidateReplacedDocument(DocX(), ItemCount, requireOpenXmlSdkConformance: false);
        WordBenchmarkCorpus.ValidateReplacedDocument(NPOI(), ItemCount);
    }

    internal void SetupOfficeAndOpenXml() {
        _fixture = WordBenchmarkCorpus.CreateParagraphFixture(ItemCount, withPlaceholder: true);
        WordBenchmarkCorpus.ValidateReplacedDocument(OfficeIMO(), ItemCount);
        WordBenchmarkCorpus.ValidateReplacedDocument(OpenXmlSdk(), ItemCount);
    }

    internal int InputBytes => _fixture.Length;

    [Benchmark(Baseline = true)]
    public byte[] OfficeIMO() {
        using var input = new MemoryStream(_fixture, writable: false);
        using WordDocument document = WordDocument.Load(input);
        document.FindAndReplace(
            WordBenchmarkCorpus.Placeholder,
            WordBenchmarkCorpus.Replacement,
            StringComparison.Ordinal);
        return document.ToBytes();
    }

    [Benchmark]
    public byte[] DocX() {
        using var input = new MemoryStream(_fixture, writable: false);
        using XceedDocX document = XceedDocX.Load(input);
        document.ReplaceText(new StringReplaceTextOptions {
            SearchValue = WordBenchmarkCorpus.Placeholder,
            NewValue = WordBenchmarkCorpus.Replacement,
            RegExOptions = RegexOptions.None
        });
        using var output = new MemoryStream();
        document.SaveAs(output);
        return output.ToArray();
    }

    [Benchmark]
    public byte[] NPOI() {
        using var input = new MemoryStream(_fixture, writable: false);
        using var document = new NpoiDocument(input);
        foreach (NPOI.XWPF.UserModel.XWPFParagraph paragraph in document.Paragraphs) {
            foreach (NPOI.XWPF.UserModel.XWPFRun run in paragraph.Runs) {
                string text = run.GetText(0) ?? string.Empty;
                if (text.Contains(WordBenchmarkCorpus.Placeholder, StringComparison.Ordinal)) {
                    run.SetText(
                        text.Replace(
                            WordBenchmarkCorpus.Placeholder,
                            WordBenchmarkCorpus.Replacement,
                            StringComparison.Ordinal),
                        0);
                }
            }
        }
        using var output = new MemoryStream();
        document.Write(output);
        return output.ToArray();
    }

    [Benchmark]
    public byte[] OpenXmlSdk() {
        using MemoryStream stream = WordBenchmarkCorpus.CreateEditableStream(_fixture);
        using (WordprocessingDocument document = WordprocessingDocument.Open(stream, isEditable: true)) {
            MainDocumentPart mainPart = document.MainDocumentPart
                ?? throw new InvalidDataException("The benchmark fixture has no main document part.");
            Body body = mainPart.Document?.Body
                ?? throw new InvalidDataException("The benchmark fixture has no body.");
            foreach (Text text in body.Descendants<Text>()) {
                text.Text = text.Text.Replace(
                    WordBenchmarkCorpus.Placeholder,
                    WordBenchmarkCorpus.Replacement,
                    StringComparison.Ordinal);
            }
            mainPart.Document!.Save();
        }
        return stream.ToArray();
    }
}
