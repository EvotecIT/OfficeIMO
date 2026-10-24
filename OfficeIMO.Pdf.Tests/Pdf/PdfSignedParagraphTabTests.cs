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
        var page = Assert.Single(pdf.GetPages());
        var first = TabMarkerStart(page, "FIELD");
        var later = TabMarkerStart(page, "LATER");
        double anchor = 30D + style.LeftIndent + position;
        Assert.InRange(first.StartBaseLine.X, anchor - .01D, anchor + .01D);
        Assert.InRange(later.StartBaseLine.X, 101.99D, 102.01D);
        Assert.True(later.StartBaseLine.Y < first.StartBaseLine.Y);
    }

    // Tight tab geometry need not produce a boundary in a reader's word heuristic.
    // These fixtures use ASCII markers; locate their painted letters directly.
    private static UglyToad.PdfPig.Content.Letter TabMarkerStart(UglyToad.PdfPig.Content.Page page, string marker) {
        var starts = Enumerable.Range(0, page.Letters.Count).Where(index =>
            string.Concat(page.Letters.Skip(index).Take(marker.Length).Select(letter => letter.Value)) == marker);
        var letters = page.Letters.Skip(Assert.Single(starts)).Take(marker.Length).ToArray();
        Assert.All(letters, letter => Assert.Equal(letters[0].StartBaseLine.Y, letter.StartBaseLine.Y, 3));
        return letters[0];
    }

    [Theory]
    [InlineData(double.NaN)]
    [InlineData(double.NegativeInfinity)]
    [InlineData(double.PositiveInfinity)]
    public void SignedTabStopsStillRejectNonfiniteCoordinates(double position) {
        Assert.Throws<ArgumentException>(() => new PdfTabStop(position));
    }
}
