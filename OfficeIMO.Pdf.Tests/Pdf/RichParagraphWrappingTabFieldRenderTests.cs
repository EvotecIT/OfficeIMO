using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class RichParagraphWrappingTests {
    [Theory]
    [InlineData(PdfTabAlignment.Left)]
    [InlineData(PdfTabAlignment.Right)]
    public void ParagraphTabs_CollapsedSpaceRetainsTabAndCompleteStyledField(PdfTabAlignment alignment) {
        const double margin = 36D, stop = 180D;
        var style = new PdfParagraphStyle { TextWhitespaceMode = PdfTextWhitespaceMode.Collapse, SpacingAfter = 0D };
        style.AddTabStop(stop, alignment, PdfTabLeaderStyle.Dots);
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
            PageWidth = 360D, PageHeight = 180D, MarginLeft = margin, MarginRight = margin,
            MarginTop = margin, MarginBottom = margin, DefaultFont = PdfStandardFont.Courier, DefaultFontSize = 12D
        }).Paragraph(p => p.Runs(new[] {
            PdfTextRun.Normal("A"), PdfTextRun.Tab(), PdfTextRun.Normal(" "),
            PdfTextRun.Normal("B "), new PdfTextRun("C", bold: true)
        }), style: style).ToBytes();

        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var page = Assert.Single(pdf.GetPages());
        var a = Assert.Single(page.Letters, letter => letter.Value == "A");
        var b = Assert.Single(page.Letters, letter => letter.Value == "B");
        var c = Assert.Single(page.Letters, letter => letter.Value == "C");
        Assert.Contains(page.Letters, letter => letter.Value == ".");
        if (alignment == PdfTabAlignment.Left) {
            double monospaceAdvance = a.EndBaseLine.X - a.StartBaseLine.X;
            Assert.True(monospaceAdvance > 0D);
            Assert.Equal(margin + stop + monospaceAdvance, b.StartBaseLine.X, 2);
        } else {
            Assert.Equal(margin + stop, c.EndBaseLine.X, 2);
        }
        Assert.Equal(a.StartBaseLine.Y, b.StartBaseLine.Y, 2);
        Assert.Equal(a.StartBaseLine.Y, c.StartBaseLine.Y, 2);
        Assert.True(c.StartBaseLine.X > b.EndBaseLine.X);
    }

    [Theory]
    [InlineData(PdfTabAlignment.Center, PdfTextWhitespaceMode.Collapse)]
    [InlineData(PdfTabAlignment.Right, PdfTextWhitespaceMode.Collapse)]
    [InlineData(PdfTabAlignment.Center, PdfTextWhitespaceMode.Preserve)]
    [InlineData(PdfTabAlignment.Right, PdfTextWhitespaceMode.Preserve)]
    public void ParagraphTabs_AlignCompleteStyledFieldsBeforeTheNextTab(PdfTabAlignment alignment, PdfTextWhitespaceMode whitespace) {
        var style = new PdfParagraphStyle { TextWhitespaceMode = whitespace, SpacingAfter = 0 };
        style.AddTabStop(180D, alignment);
        style.AddTabStop(300D);
        byte[] bytes = PdfDocument.Create(new PdfCreateOptions {
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
