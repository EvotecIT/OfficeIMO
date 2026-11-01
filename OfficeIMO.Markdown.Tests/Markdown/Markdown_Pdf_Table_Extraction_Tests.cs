using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.MarkdownSuite;

public sealed class Markdown_Pdf_Table_Extraction_Tests {
    private const string Header = "| Item | Qty | Total |\n|---|---|---|\n| Alpha | 1 | 10.00 |\n| Beta | 2 | 20.00 |\n| Gamma | 3 | 30.00 |\n";

    [Fact]
    public void TrailingRowWithOnlyATotalIsExtracted() {
        string[][] rows = ExtractRows(Header + "| | | 100.00 |\n");

        Assert.Equal(new[] { string.Empty, string.Empty, "100.00" }, rows[rows.Length - 1]);
        Assert.Equal(5, rows.Length);
    }

    [Fact]
    public void MiddleSparseRowStaysSeparateFromTheRowAboveIt() {
        string[][] rows = ExtractRows(Header + "| | 4 | |\n| Delta | 5 | 50.00 |\n");

        Assert.Equal(new[] { "Gamma", "3", "30.00" }, rows[3]);
        Assert.Equal(new[] { string.Empty, "4", string.Empty }, rows[4]);
        Assert.Equal(new[] { "Delta", "5", "50.00" }, rows[5]);
    }

    private static string[][] ExtractRows(string markdown) {
        byte[] pdf = MarkdownDocument.Parse(markdown).ToPdf().ToBytes();
        PdfStructuredTablePage page = Assert.Single(PdfTextExtractor.ExtractTablesByPage(pdf));
        return Assert.Single(page.Tables).Rows.ToArray();
    }
}
