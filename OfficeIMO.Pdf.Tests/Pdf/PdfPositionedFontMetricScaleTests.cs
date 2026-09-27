using System;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfPositionedFontMetricScaleTests {
    [Theory]
    [InlineData(false, OfficeTextDirection.Auto)]
    [InlineData(false, OfficeTextDirection.LeftToRight)]
    [InlineData(false, OfficeTextDirection.RightToLeft)]
    [InlineData(true, OfficeTextDirection.Auto)]
    [InlineData(true, OfficeTextDirection.LeftToRight)]
    [InlineData(true, OfficeTextDirection.RightToLeft)]
    public void PositionedMetricScaleKeepsNativePointTextIndependent(bool named, OfficeTextDirection direction) {
        byte[] data = ManagedTextShapingTestAssets.AddTrackingTable(
            ManagedTextShapingTestAssets.CreateFontWithDistinctGlyphs(' ', 'A', 'B'),
            ManagedTextShapingTestAssets.CreateTrackingTable(new[] { 10D, 20D }, new[] { 0D },
                new[] { new short[] { -100, 0 } }));
        var options = new PdfOptions { CompressContentStreams = false };
        if (named) options.RegisterNamedFontFamily(new PdfEmbeddedFontFamily("Tracked", data));
        else options.EmbedStandardFont(PdfStandardFont.Helvetica, data);
        PdfTextRun run = PdfTextRun.Normal("AB", fontSize: 15D, fontFamily: named ? "Tracked" : null)
            .WithTextDirection(direction);
        byte[] pdf = PdfDocument.Create(document => document.Page(page => page.Canvas(canvas => {
            canvas.PositionedText(new[] { run }, PdfCanvasTextStructureRole.Paragraph,
                10D, 10D, 100D, 30D, fontSize: 15D, advanceWidth: 15D, fontMetricScale: .75D);
            canvas.PositionedText(new[] { run }, PdfCanvasTextStructureRole.Paragraph,
                10D, 50D, 100D, 30D, fontSize: 15D, advanceWidth: 13.5D);
        })), options).ToBytes();
        string content = Encoding.ASCII.GetString(pdf);
        Assert.Contains("<0002>] TJ", content, StringComparison.Ordinal);
        Assert.Contains("<0002> 50] TJ", content, StringComparison.Ordinal);
        Assert.DoesNotContain("1.111111", content, StringComparison.Ordinal);
        Assert.Contains("AB", PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Fact]
    public void TrackedWidthUsesAuthoredSizeAndPhysicalGeometry() {
        var font = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateTrackingFont(
            PdfFontTrackingTests.Table()));
        PdfGlyphRun run = font.ShapeText("AB");
        Assert.Equal(15D, font.MeasureShapedTextWidth("AB", run, 15D, .75D), 9);
        Assert.Equal(13.5D, font.MeasureShapedTextWidth("AB", run, 15D), 9);
    }

    [Theory]
    [InlineData(false, 40)]
    [InlineData(true, 960)]
    public void AuthoredSizeRetainsSignedClusterAndAttachmentSemantics(bool negative, int postAdjustment) {
        var font = PdfTrueTypeFontProgram.Parse(ManagedTextShapingTestAssets.CreateTrackingFont(
            PdfFontTrackingTests.Table()));
        var run = new PdfGlyphRun(new[] {
            new PdfGlyphInfo(1, "A", 0, 500, negative ? -500 : 500, 20, 0),
            new PdfGlyphInfo(1, "\u0301", 1, 500, 0, -30, 50),
            new PdfGlyphInfo(1, "B", 2, 500, negative ? -500 : 500, 0, 0)
        }, Array.Empty<PdfTextEncodingDiagnostic>(), "A\u0301B",
            negative ? OfficeTextDirection.RightToLeft : OfficeTextDirection.LeftToRight);
        var content = new StringBuilder();
        new ContentStreamBuilder(content).ShowText(font.ToTextShowCommand("A\u0301B", run, .75D), 12D);
        Assert.Contains("[<0001> " + postAdjustment + "] TJ", content.ToString(), StringComparison.Ordinal);
        Assert.Contains(negative ? "[-20 <0001> 1020] TJ" : "[-20 <0001> 20] TJ", content.ToString(), StringComparison.Ordinal);
        Assert.Equal(11.04D, font.MeasureShapedTextWidth("A\u0301B", run, 12D, .75D), 9);
    }
}
