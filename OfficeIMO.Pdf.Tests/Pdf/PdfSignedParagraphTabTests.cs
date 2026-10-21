using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfSignedParagraphTabTests {
    [Theory]
    [InlineData(-25D)]
    [InlineData(0D)]
    public void SignedTabStopsApplyAheadOfAHangingFirstLineAndAreSkippedOnLaterLines(double position) {
        var style = new PdfParagraphStyle {
            LeftIndent = 36D, FirstLineIndent = -36D, DefaultTabStopWidth = 36D, SpacingAfter = 0D
        };
        style.AddTabStop(position);
        byte[] bytes = PdfDocument.Create(new PdfOptions {
            PageWidth = 300D, PageHeight = 180D, MarginLeft = 30D, MarginRight = 30D,
            MarginTop = 30D, MarginBottom = 30D, DefaultFontSize = 10D
        }).Paragraph(paragraph => paragraph.Text("1.\tFIELD\nX\tLATER"), style: style).ToBytes();
        using var pdf = UglyToad.PdfPig.PdfDocument.Open(bytes);
        var words = Assert.Single(pdf.GetPages()).GetWords().ToArray();
        var first = Assert.Single(words, word => word.Text == "FIELD");
        var later = Assert.Single(words, word => word.Text == "LATER");
        double anchor = 30D + style.LeftIndent + position;
        Assert.InRange(first.Letters.First().StartBaseLine.X, anchor - .01D, anchor + .01D);
        Assert.InRange(later.Letters.First().StartBaseLine.X, 101.99D, 102.01D);
        Assert.True(later.Letters.First().StartBaseLine.Y < first.Letters.First().StartBaseLine.Y);
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.NegativeInfinity)]
    [InlineData(double.PositiveInfinity)]
    public void SignedTabStopsStillRejectNonfiniteCoordinates(double position) {
        Assert.Throws<ArgumentException>(() => new PdfTabStop(position));
    }
}
