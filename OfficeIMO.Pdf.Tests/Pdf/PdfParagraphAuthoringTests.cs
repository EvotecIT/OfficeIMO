using OfficeIMO.Pdf;
using System.Linq;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfParagraphAuthoringTests {
    [Fact]
    public void AddText_SetBoldFormatsOnlyThatRun_AndRunDefaultsApplyToLaterRuns() {
        var document = PdfDocument.Create();
        PdfParagraph paragraph = document.AddParagraph();
        paragraph.AddText("BoldRun ").SetBold();
        paragraph.AddText("PlainRun ");
        paragraph.RunDefaults.Italic = true;
        paragraph.AddText("ItalicRun");

        var spans = PdfReadDocument.Open(document.ToBytes()).Pages.SelectMany(page => page.GetTextSpans()).ToArray();

        var bold = Assert.Single(spans, span => span.Text.Contains("BoldRun"));
        Assert.True(bold.IsBold);
        Assert.False(bold.IsItalic);
        var plain = Assert.Single(spans, span => span.Text.Contains("PlainRun"));
        Assert.False(plain.IsBold);
        Assert.False(plain.IsItalic);
        var italic = Assert.Single(spans, span => span.Text.Contains("ItalicRun"));
        Assert.True(italic.IsItalic);
        Assert.False(italic.IsBold);
    }

    [Fact]
    public void AddParagraph_ReturnsLiveParagraph_WhoseLaterEditsAreRendered() {
        var document = PdfDocument.Create();
        PdfParagraph paragraph = document.AddParagraph("FirstRun ");
        document.AddParagraph("NextParagraph");
        paragraph.AddText("LateRun").SetBold();
        paragraph.Runs[0].SetItalic();

        var spans = PdfReadDocument.Open(document.ToBytes()).Pages.SelectMany(page => page.GetTextSpans()).ToArray();

        Assert.True(Assert.Single(spans, span => span.Text.Contains("FirstRun")).IsItalic);
        var late = Assert.Single(spans, span => span.Text.Contains("LateRun"));
        Assert.True(late.IsBold);
        Assert.True(late.Y > Assert.Single(spans, span => span.Text.Contains("NextParagraph")).Y);
    }
}
