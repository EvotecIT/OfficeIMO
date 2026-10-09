using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PdfMixedInlineImageUnavailableTests {
    [Theory]
    [InlineData("body")]
    [InlineData("cell")]
    [InlineData("header")]
    public void LinkedInlineImageLeavesTextAndAnUnavailableImageWarning(string story) {
        using WordDocument word = WordDocument.Create();
        word.AddParagraph("Body");
        if (story == "header") word.AddHeadersAndFooters();
        var paragraph = story == "cell" ? word.AddTable(1, 1).Rows[0].Cells[0].Paragraphs[0] :
            story == "header" ? word.Sections[0].Header.Default!.AddParagraph() : word.AddParagraph();
        paragraph.Text = "Before";
        paragraph.AddText("").AddImage(new Uri("https://example.invalid/image.png"), 80, 80);
        paragraph.AddText("After");
        var result = word.ToPdfDocumentResult();
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "NativeBodyImageUnavailable");
        var read = PdfReadDocument.Open(result.Value.ToBytes());
        Assert.Contains("BeforeAfter", read.ExtractText());
        Assert.Empty(read.ExtractImages());
    }
}
