using OfficeIMO.Word;
using OfficeIMO.Word.Pdf;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PdfMixedInlineImageUnavailableTests {
    [Fact]
    public void LinkedInlineImageLeavesTextAndAnUnavailableImageWarning() {
        using WordDocument word = WordDocument.Create();
        var paragraph = word.AddParagraph("Before");
        paragraph.AddText("").AddImage(new Uri("https://example.invalid/image.png"), 80, 80);
        paragraph.AddText("After");
        var result = word.ToPdfDocumentResult();
        Assert.Contains(result.Report.Warnings, warning => warning.Code == "NativeBodyImageUnavailable");
        var read = PdfReadDocument.Open(result.Value.ToBytes());
        Assert.Contains("BeforeAfter", read.ExtractText());
        Assert.Empty(read.ExtractImages());
    }
}
