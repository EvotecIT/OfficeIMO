using BenchmarkDotNet.Attributes;
using OfficeIMO.Word.Html;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>
/// Reads a Word-authored report whose styles part (about 300 KB of style, table-style and latent-style
/// definitions) dwarfs its body, the case where reading styles costs more than the content that uses them.
/// Each benchmark loads the package read-only and then finds a few styles, lists every style, inspects the
/// document, or converts it to Markdown or HTML.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Word", "Styles")]
public class WordStyleCatalogBenchmarks {
    private const string FixtureName = "native-word-report.docx";
    private byte[] _document = Array.Empty<byte>();

    [GlobalSetup]
    public void Setup() {
        _document = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, FixtureName));
        if (ListStyles() < 100) throw new InvalidOperationException("The style-heavy fixture lists fewer styles than expected.");
        if (FindStyles() == 0) throw new InvalidOperationException("The style-heavy fixture has no styles to find.");
        if (Inspect() == 0) throw new InvalidOperationException("Inspection found no paragraphs.");
        if (ToMarkdown() == 0 || ToHtml() == 0) throw new InvalidOperationException("A conversion produced no output.");
    }

    [Benchmark]
    public int FindStyles() {
        using WordDocument document = Load();
        int length = document.Styles.Count;
        foreach (string styleId in new[] { "Normal", "Heading1", "Heading2", "TableGrid" }) length += document.Styles[styleId].Name.Length;
        return length;
    }

    [Benchmark]
    public int ListStyles() {
        using WordDocument document = Load();
        int length = 0;
        foreach (WordStyle style in document.Styles) length += style.Name.Length;
        return length;
    }

    [Benchmark]
    public int Inspect() {
        using WordDocument document = Load();
        return document.Inspect(includeImageBytes: false).Sections.Sum(section => section.Elements.Count);
    }

    [Benchmark]
    public int ToMarkdown() {
        using WordDocument document = Load();
        return document.ToMarkdown().ToMarkdown().Length;
    }

    [Benchmark]
    public int ToHtml() {
        using WordDocument document = Load();
        return document.ToHtml().Length;
    }

    private WordDocument Load() =>
        WordDocument.Load(new MemoryStream(_document, writable: false), new WordLoadOptions { AccessMode = DocumentAccessMode.ReadOnly });
}
