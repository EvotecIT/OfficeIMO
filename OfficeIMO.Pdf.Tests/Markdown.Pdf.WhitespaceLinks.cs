using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using PdfCore = OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class MarkdownPdfTests {
    [Fact]
    public void Markdown_SaveAsPdf_UsesDestinationForWhitespaceOnlyLinkLabels() {
        const string paragraphUrl = "https://example.com/paragraph";
        const string tableUrl = "https://example.com/table";
        string markdown = """
Paragraph [ ](https://example.com/paragraph " ").

| Surface | Link |
| --- | --- |
| Table | [ ](https://example.com/table " ") |
""";

        byte[] pdf = MarkdownReader.Parse(markdown).ToPdfBytes();
        PdfCore.PdfDocumentReadResult logical = PdfCore.PdfDocumentReadResult.Load(pdf);
        string text = PdfCore.PdfReadDocument.Open(pdf).ExtractText();

        Assert.Contains(paragraphUrl, text);
        Assert.Contains(tableUrl, text);
        Assert.Equal(paragraphUrl, Assert.Single(logical.GetLinksByUri(paragraphUrl)).Contents);
        Assert.Equal(tableUrl, Assert.Single(logical.GetLinksByUri(tableUrl)).Contents);
    }
}
