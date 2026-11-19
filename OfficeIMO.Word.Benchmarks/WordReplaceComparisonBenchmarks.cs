using System.Text.RegularExpressions;
using OfficeIMO.Benchmarks;
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
        if (BenchmarkLibrarySelection.IsEnabled("docx")) WordBenchmarkCorpus.Check("DocX", "Default", Content,
            () => WordBenchmarkCorpus.ValidateReplacedDocument(DocX(), ItemCount, requireOpenXmlSdkConformance: false));
        byte[] npoi = NPOI();
        WordBenchmarkCorpus.Check("NPOI", "Default", Conformant,
            () => WordBenchmarkCorpus.ValidateReplacedDocument(npoi, ItemCount));
#if NET10_0_OR_GREATER
        byte[] oneDocument = OneDocument();
        WordBenchmarkCorpus.Check("OneDocument.Documents", "Default", Conformant,
            () => WordBenchmarkCorpus.ValidateReplacedDocument(oneDocument, ItemCount));
#endif
    }

    internal void SetupOfficeAndOpenXml() {
        _fixture = WordBenchmarkCorpus.CreateParagraphFixture(ItemCount, withPlaceholder: true);
        byte[] officeImo = OfficeIMO();
        WordBenchmarkCorpus.Check("OfficeIMO", "Default", Conformant,
            () => WordBenchmarkCorpus.ValidateReplacedDocument(officeImo, ItemCount));
        byte[] openXmlSdk = OpenXmlSdk();
        WordBenchmarkCorpus.Check("Open XML SDK", "Default", Conformant,
            () => WordBenchmarkCorpus.ValidateReplacedDocument(openXmlSdk, ItemCount));
    }

    // The saved output is reopened as a ZIP/XML package and every paragraph must carry the replaced text.
    private string Content => $"ItemCount={ItemCount}; {ItemCount} paragraphs with the placeholder replaced";

    private string Conformant => Content + "; passes the Open XML SDK validator";

    internal int InputBytes => _fixture.Length;

    [Benchmark(Baseline = true)]
    [BenchmarkCategory("library:OfficeIMO", "case:Default")]
    public byte[] OfficeIMO() {
        using var input = new MemoryStream(_fixture, writable: false);
        using WordDocument document = WordDocument.Load(input);
        document.ReplaceAll(
            WordBenchmarkCorpus.Placeholder,
            WordBenchmarkCorpus.Replacement,
            new global::OfficeIMO.OfficeTextSearchOptions { MatchCase = true });
        return document.ToBytes();
    }

    [Benchmark]
    [BenchmarkCategory("library:DocX", "case:Default")]
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
    [BenchmarkCategory("library:NPOI", "case:Default")]
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
                            WordBenchmarkCorpus.Replacement),
                        0);
                }
            }
        }
        using var output = new MemoryStream();
        document.Write(output);
        return output.ToArray();
    }

#if NET10_0_OR_GREATER
    [Benchmark]
    [BenchmarkCategory("library:OneDocument.Documents", "case:Default")]
    public byte[] OneDocument() => WordOneDocumentWorkloads.Replace(_fixture);

#endif
    [Benchmark]
    [BenchmarkCategory("library:Open XML SDK", "case:Default")]
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
                    WordBenchmarkCorpus.Replacement);
            }
            mainPart.Document!.Save();
        }
        return stream.ToArray();
    }
}
