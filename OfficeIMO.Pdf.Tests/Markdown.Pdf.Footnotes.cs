using OfficeIMO.Markdown;
using OfficeIMO.Markdown.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public partial class MarkdownPdfTests {
    [Fact]
    public void StructuredFootnoteRendersEachParagraphOnceAndKeepsInlineFormatting() {
        var body = MarkdownReader.Parse("First **bold** paragraph.\n\nSecond distinct paragraph.\n\n- Alpha item\n- Beta item\n\n> Quoted note.");
        var document = MarkdownDoc.Create().Add(new FootnoteDefinitionBlock("note", body.Blocks));
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(document.ToPdfBytes()).ExtractText();
        foreach (string marker in new[] { "First", "Second", "Alpha", "Beta", "Quoted" })
            Assert.Single(System.Text.RegularExpressions.Regex.Matches(text, marker).Cast<System.Text.RegularExpressions.Match>());
        Assert.DoesNotContain("**bold**", text);
    }

    [Theory]
    [InlineData("- First list item\n- Second list item")]
    [InlineData("> First quoted paragraph.")]
    public void StructuredFootnoteStartingWithNonParagraphKeepsItsFirstBlock(string markdown) {
        var body = MarkdownReader.Parse(markdown);
        var document = MarkdownDoc.Create().Add(new FootnoteDefinitionBlock("note", body.Blocks));
        string text = OfficeIMO.Pdf.PdfReadDocument.Open(document.ToPdfBytes()).ExtractText();
        Assert.Single(System.Text.RegularExpressions.Regex.Matches(text, "First").Cast<System.Text.RegularExpressions.Match>());
        Assert.Contains("note", text);
        Assert.DoesNotContain("> First", text);
        Assert.DoesNotContain("- First", text);
    }
}
