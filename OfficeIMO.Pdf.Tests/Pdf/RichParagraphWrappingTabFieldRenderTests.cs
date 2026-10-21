using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class RichParagraphWrappingTests {
    [Theory]
    [InlineData(PdfTabAlignment.Center, PdfTextWhitespaceMode.Collapse)]
    [InlineData(PdfTabAlignment.Right, PdfTextWhitespaceMode.Collapse)]
    [InlineData(PdfTabAlignment.Center, PdfTextWhitespaceMode.Preserve)]
    [InlineData(PdfTabAlignment.Right, PdfTextWhitespaceMode.Preserve)]
    public void ParagraphTabs_AlignCompleteStyledFieldsBeforeTheNextTab(PdfTabAlignment alignment, PdfTextWhitespaceMode whitespace) {
        var style = new PdfParagraphStyle { TextWhitespaceMode = whitespace, SpacingAfter = 0 };
        style.AddTabStop(180D, alignment);
        style.AddTabStop(300D);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
            PageWidth = 400D, PageHeight = 180D, MarginLeft = 40D, MarginRight = 40D,
            MarginTop = 40D, MarginBottom = 40D, DefaultFontSize = 10D
        }).Paragraph(p => p.Runs(new[] {
            PdfTextRun.Normal("LABEL\tFooter "), new PdfTextRun("Right", bold: true),
            PdfTextRun.Normal("\tTAIL")
        }), style: style).ToBytes();

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var words = Assert.Single(pdf.GetPages()).GetWords().ToArray();
        var label = Assert.Single(words, word => word.Text == "LABEL");
        var footer = Assert.Single(words, word => word.Text == "Footer");
        var right = Assert.Single(words, word => word.Text == "Right");
        var tail = Assert.Single(words, word => word.Text == "TAIL");
        double anchor = alignment == PdfTabAlignment.Center
            ? (footer.BoundingBox.Left + right.BoundingBox.Right) / 2D : right.BoundingBox.Right;
        Assert.InRange(anchor, 219D, 221D);
        Assert.InRange(tail.BoundingBox.Left, 339D, 341D);
        foreach (var word in new[] { footer, right, tail })
            Assert.Equal(label.Letters[0].StartBaseLine.Y, word.Letters[0].StartBaseLine.Y, 2);
    }
}
