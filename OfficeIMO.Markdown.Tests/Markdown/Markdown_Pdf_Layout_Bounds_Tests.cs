using System.Text;
using OfficeIMO.Core;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_Pdf_Layout_Bounds_Tests {
    [Theory]
    [InlineData("```\n```\n")]
    [InlineData("```\n\n```\n")]
    public void BlankCodeFenceExports(string markdown) {
        Assert.NotEmpty(MarkdownDocument.Parse(markdown).ToPdf().ToBytes());
    }

    [Theory]
    [InlineData("x", 4_000)]
    [InlineData("x", 20_000)]
    [InlineData("q ", 10_000)]
    public void ShortCodeFenceTallerThanAPageContinuesOnTheNextPage(string text, int repeats) {
        string line = string.Concat(Enumerable.Repeat(text, repeats));

        byte[] pdf = MarkdownDocument.Parse("```\n" + line + "\n```\n").ToPdf().ToBytes();

        Assert.True(PdfTextExtractor.ExtractTextByPage(pdf).Count > 1);
        Assert.Equal(repeats, PdfTextExtractor.ExtractAllText(pdf).Count(ch => ch == text[0]));
    }

    [Fact]
    public void ShortCodeFenceMovesWholeToTheNextPage() {
        var markdown = new StringBuilder();
        for (int index = 0; index < 34; index++) {
            markdown.Append("Filler ").Append(index).Append("\n\n");
        }
        markdown.Append("```\n");
        for (int index = 0; index < 15; index++) {
            markdown.Append("listing-").Append(index).Append('\n');
        }
        markdown.Append("```\n");

        IReadOnlyList<string> pages = PdfTextExtractor.ExtractTextByPage(MarkdownDocument.Parse(markdown.ToString()).ToPdf().ToBytes());

        Assert.True(pages.Count > 1);
        string page = Assert.Single(pages, text => text.Contains("listing-0", StringComparison.Ordinal));
        Assert.Contains("listing-14", page, StringComparison.Ordinal);
    }

    [Fact]
    public void ListNestedTenLevelsKeepsFullIndentationOnTheLeftOnly() {
        (byte[] pdf, OfficeConversionReport report) = Export(BuildNestedList(10));

        Assert.Contains("level-9", PdfTextExtractor.ExtractAllText(pdf), StringComparison.Ordinal);
        Assert.DoesNotContain(report.Diagnostics, diagnostic => diagnostic.Code == "panel-indent-reduced");
    }

    [Fact]
    public void ListNestedBeyondThePageWidthClampsItsIndentationWithAWarning() {
        (byte[] pdf, OfficeConversionReport report) = Export(BuildNestedList(40));

        Assert.Contains("level-39", WithoutWhitespace(PdfTextExtractor.ExtractAllText(pdf)), StringComparison.Ordinal);
        Assert.Single(report.Diagnostics, diagnostic => diagnostic.Code == "panel-indent-reduced");
    }

    [Fact]
    public void QuoteNestedThirtyLevelsClampsItsPaddingWithAWarning() {
        (byte[] pdf, OfficeConversionReport report) = Export(new string('>', 30) + " deep-quote");

        Assert.Contains("deep-quote", PdfTextExtractor.ExtractAllText(pdf), StringComparison.Ordinal);
        Assert.Single(report.Diagnostics, diagnostic => diagnostic.Code == "panel-indent-reduced");
    }

    [Fact]
    public void TableWiderThanItsCellPaddingAllowsDropsThePaddingWithAWarning() {
        const int columns = 60;
        string markdown = "|" + string.Join("|", Enumerable.Range(0, columns).Select(index => "h" + index)) + "|\n"
            + "|" + string.Concat(Enumerable.Repeat("---|", columns)) + "\n"
            + "|" + string.Join("|", Enumerable.Range(0, columns).Select(index => "c" + index)) + "|\n";

        (byte[] pdf, OfficeConversionReport report) = Export(markdown);

        Assert.NotEmpty(PdfTextExtractor.ExtractAllText(pdf));
        Assert.Single(report.Diagnostics, diagnostic => diagnostic.Code == "table-cell-padding-dropped");
    }

    private static (byte[] Pdf, OfficeConversionReport Report) Export(string markdown) {
        OfficeIMO.Pdf.PdfDocument document = MarkdownDocument.Parse(markdown).ToPdf(new MarkdownPdfExportOptions(), out OfficeConversionReport report);
        return (document.ToBytes(), report);
    }

    private static string WithoutWhitespace(string text) =>
        string.Concat(text.Where(ch => !char.IsWhiteSpace(ch)));

    private static string BuildNestedList(int levels) {
        var markdown = new StringBuilder();
        for (int level = 0; level < levels; level++) {
            markdown.Append(' ', level * 2).Append("- level-").Append(level).Append('\n');
        }

        return markdown.ToString();
    }
}
