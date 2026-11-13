using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Markdown;
using OfficeIMO.Word.Markdown;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>
/// Measures Markdown-to-Word and Word-to-Markdown conversion at two document sizes, so
/// time and allocation growth between the sizes shows whether either direction scales linearly.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Word", "Markdown")]
public class WordMarkdownConversionBenchmarks {
    private string _markdown = string.Empty;
    private byte[] _document = Array.Empty<byte>();

    [Params(100, 400)]
    public int SectionCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        var builder = new StringBuilder();
        for (int index = 0; index < SectionCount; index++) {
            string number = index.ToString(CultureInfo.InvariantCulture);
            builder.Append("## Section ").Append(number).Append("\n\n");
            builder.Append("Paragraph ").Append(number).Append(" with **bold** and `code` text.\n\n");
            builder.Append("- First item\n- Second item\n  - Nested item\n\n");
            builder.Append("| Name | Value |\n| --- | --- |\n| Row ").Append(number).Append(" | Cell |\n| Other | Cell |\n\n");
        }
        _markdown = builder.ToString();
        _document = ConvertMarkdownToWord();

        Validate();
    }

    [Benchmark]
    public int MarkdownToWord() => ConvertMarkdownToWord().Length;

    [Benchmark]
    public int WordToMarkdown() => ConvertWordToMarkdown().Length;

    private byte[] ConvertMarkdownToWord() {
        using WordDocument document = MarkdownDocument.Parse(_markdown).ToWord();
        return document.ToBytes();
    }

    private string ConvertWordToMarkdown() {
        using var stream = new MemoryStream(_document, writable: false);
        using WordDocument document = WordDocument.Load(stream);
        return document.ToMarkdown().ToMarkdown();
    }

    private void Validate() {
        using (var stream = new MemoryStream(_document, writable: false))
        using (WordDocument document = WordDocument.Load(stream)) {
            EnsureEqual("Word tables", SectionCount, document.Tables.Count);
        }

        string markdown = ConvertWordToMarkdown();
        int headings = 0;
        int rows = 0;
        foreach (string line in markdown.Split('\n')) {
            if (line.StartsWith("## Section ", StringComparison.Ordinal)) headings++;
            if (line.StartsWith("| Row ", StringComparison.Ordinal)) rows++;
        }
        EnsureEqual("Markdown headings", SectionCount, headings);
        EnsureEqual("Markdown table rows", SectionCount, rows);
    }

    private static void EnsureEqual(string what, int expected, int actual) {
        if (actual != expected) {
            throw new InvalidOperationException(what + " returned " + actual + "; expected " + expected + ".");
        }
    }
}
