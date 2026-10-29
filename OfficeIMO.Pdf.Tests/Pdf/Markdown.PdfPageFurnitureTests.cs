using System;
using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using Xunit;
using PdfCore = OfficeIMO.Pdf;

namespace OfficeIMO.Tests.Pdf;

public class MarkdownPdfPageFurnitureTests {
    private const string Settings = "# Settings\n\n| Setting | Type | Default | Description |\n|:--|:--|--:|:--|\n" +
        "| `search.maxResults` | integer | `20` | Maximum number of hits returned per query, which is a long description that wraps. |\n" +
        "| `title` | string | `\"Lumen site\"` | Site title shown in the header and browser tab of every generated page. |\n";

    [Fact]
    public void DefaultExportNumbersPagesUnlessTurnedOff() {
        string numbered = PdfCore.PdfReadDocument.Open(MarkdownDocument.Parse("# Title\n\nBody").ToPdf(new MarkdownPdfExportOptions()).ToBytes()).ExtractText();
        string plain = PdfCore.PdfReadDocument.Open(MarkdownDocument.Parse("# Title\n\nBody")
            .ToPdf(new MarkdownPdfExportOptions { ShowPageNumbers = false }).ToBytes()).ExtractText();

        Assert.Contains("Page 1 of 1", numbered, StringComparison.Ordinal);
        Assert.DoesNotContain("Page 1 of 1", plain, StringComparison.Ordinal);
    }

    [Fact]
    public void InlineCodeIdentifiersAndShortValuesStayWholeInTableCells() {
        string text = PdfCore.PdfReadDocument.Open(MarkdownDocument.Parse(Settings).ToPdf(new MarkdownPdfExportOptions()).ToBytes()).ExtractText();

        Assert.Contains("search.maxResults", text, StringComparison.Ordinal);
        Assert.Contains("\"Lumen site\"", text, StringComparison.Ordinal);
    }
}
