using System.Globalization;
using System.Text;
using BenchmarkDotNet.Attributes;
using OfficeIMO.Html;
using OfficeIMO.Word.Html;

namespace OfficeIMO.Word.Benchmarks;

/// <summary>
/// Measures HTML-to-Word conversion for container-heavy, table-heavy and list-heavy input at two
/// sizes, so time and allocation growth between the sizes shows whether the import scales linearly.
/// </summary>
[MemoryDiagnoser]
[BenchmarkCategory("Word", "Html")]
public class WordHtmlImportBenchmarks {
    private string _html = string.Empty;

    [Params(HtmlImportShape.Containers, HtmlImportShape.Tables, HtmlImportShape.Lists)]
    public HtmlImportShape Shape { get; set; }

    [Params(100, 400)]
    public int BlockCount { get; set; }

    [GlobalSetup]
    public void Setup() {
        _html = BuildHtml(Shape, BlockCount);
        Validate();
    }

    [Benchmark]
    public int HtmlToWord() => ConvertHtmlToWord().Length;

    private byte[] ConvertHtmlToWord() {
        using WordDocument document = HtmlConversionDocument.Parse(_html).ToWord();
        return document.ToBytes();
    }

    internal static string BuildHtml(HtmlImportShape shape, int blockCount) {
        var builder = new StringBuilder("<html><body>");
        for (int index = 0; index < blockCount; index++) {
            string number = index.ToString(CultureInfo.InvariantCulture);
            switch (shape) {
                case HtmlImportShape.Containers:
                    builder.Append("<div class=\"block\"><h2>Block ").Append(number).Append("</h2>");
                    builder.Append("<div style=\"margin-left:12pt\"><p>Paragraph ").Append(number)
                        .Append(" with <b>bold</b> text.</p>Loose text <i>inline</i>");
                    builder.Append("<article><p>Article text</p></article></div>");
                    builder.Append("<blockquote><p>Quote ").Append(number).Append("</p></blockquote></div>");
                    break;
                case HtmlImportShape.Tables:
                    builder.Append("<div><p>Table ").Append(number).Append("</p>");
                    builder.Append("<table><tr><th>Name</th><th>Value</th></tr><tr><td>Row ").Append(number)
                        .Append("</td><td><div><p>Cell</p></div></td></tr><tr><td>Other</td><td>Cell</td></tr></table></div>");
                    break;
                case HtmlImportShape.Lists:
                    builder.Append("<div><p>List ").Append(number).Append("</p>");
                    builder.Append("<ul><li>First item</li><li>Second item<ol><li>Nested item</li></ol></li></ul></div>");
                    break;
            }
        }
        return builder.Append("</body></html>").ToString();
    }

    private void Validate() {
        using var stream = new MemoryStream(ConvertHtmlToWord(), writable: false);
        using WordDocument document = WordDocument.Load(stream);
        string expectedPrefix = Shape switch {
            HtmlImportShape.Containers => "Block ",
            HtmlImportShape.Tables => "Table ",
            _ => "List "
        };
        int markers = document.Paragraphs.Count(paragraph => paragraph.Text.StartsWith(expectedPrefix, StringComparison.Ordinal));
        EnsureEqual("Block marker paragraphs", BlockCount, markers);
        if (Shape == HtmlImportShape.Tables) {
            EnsureEqual("Word tables", BlockCount, document.Tables.Count);
        } else if (Shape == HtmlImportShape.Lists) {
            EnsureEqual("List items", BlockCount * 3, document.Paragraphs.Count(paragraph => paragraph.IsListItem && paragraph.Text.EndsWith(" item", StringComparison.Ordinal)));
        } else {
            EnsureEqual("Quote paragraphs", BlockCount, document.Paragraphs.Count(paragraph => paragraph.Text.StartsWith("Quote ", StringComparison.Ordinal)));
        }
    }

    private static void EnsureEqual(string what, int expected, int actual) {
        if (actual != expected) {
            throw new InvalidOperationException(what + " returned " + actual + "; expected " + expected + ".");
        }
    }
}

/// <summary>Kinds of generated HTML input measured by <see cref="WordHtmlImportBenchmarks"/>.</summary>
public enum HtmlImportShape {
    /// <summary>Nested div, article and blockquote containers around paragraphs and loose inline text.</summary>
    Containers,
    /// <summary>Small tables, each preceded by a paragraph inside a div, with a nested div in one cell.</summary>
    Tables,
    /// <summary>Nested bullet and numbered lists inside divs.</summary>
    Lists
}
