using System;
using System.IO;
using System.Linq;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;
using PdfPigDocument = UglyToad.PdfPig.PdfDocument;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfClippedDrawingGroupTests {
    [Fact]
    public void DrawingFlowRendersClippedNestedGroups() {
        var child = new OfficeDrawing(80D, 40D)
            .AddShape(new OfficeShape {
                Kind = OfficeShapeKind.Rectangle,
                Width = 80D,
                Height = 40D,
                FillColor = OfficeColor.FromRgb(37, 99, 235)
            }, 0D, 0D);
        var drawing = new OfficeDrawing(60D, 40D)
            .AddClippedDrawing(
                child,
                10D,
                5D,
                OfficeClipPath.RoundedRectangle(40D, 25D, 4D),
                contentOffsetX: -12D,
                contentOffsetY: -3D);

        byte[] pdf = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Drawing(drawing)
            .ToBytes();
        string content = Encoding.ASCII.GetString(pdf);

        Assert.Contains(" W n", content);
        Assert.Equal("%PDF", Encoding.ASCII.GetString(pdf, 0, 4));
    }

    [Fact]
    public void NonWrappingDrawingTextStaysOnOneBaselineInsideClippedGroups() {
        var child = new OfficeDrawing(40D, 30D)
            .AddText(
                "OK",
                10.08D,
                3D,
                19.84D,
                20D,
                font: new OfficeFontInfo("Arial", 16D),
                color: OfficeColor.White,
                wrapText: false);
        var drawing = new OfficeDrawing(160D, 100D)
            .AddClippedDrawing(child, 0D, 0D, OfficeClipPath.Rectangle(40D, 30D));

        byte[] pdf = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Drawing(drawing)
            .ToBytes();
        using PdfPigDocument parsed = PdfPigDocument.Open(new MemoryStream(pdf));
        var letters = parsed.GetPage(1).Letters
            .Where(letter => letter.Value == "O" || letter.Value == "K")
            .ToList();

        Assert.Equal(2, letters.Count);
        Assert.InRange(Math.Abs(letters[0].StartBaseLine.Y - letters[1].StartBaseLine.Y), 0D, 0.01D);
        Assert.True(letters[1].StartBaseLine.X > letters[0].StartBaseLine.X);
    }

    [Fact]
    public void DrawingFlowRetainsActualTextForVectorGlyphPaint() {
        OfficeShape glyphPaint = OfficeShape.Rectangle(20D, 12D);
        glyphPaint.FillColor = OfficeColor.Navy;
        var paint = new OfficeDrawing(40D, 20D).AddShape(glyphPaint, 4D, 3D);
        var drawing = new OfficeDrawing(40D, 20D)
            .AddActualTextDrawing("Vector label", paint, 4D, 15D);

        OfficeDrawing clone = drawing.Clone();
        OfficeDrawingGroup logicalGroup = Assert.Single(clone.Elements.OfType<OfficeDrawingGroup>());
        Assert.Equal("Vector label", logicalGroup.ActualText);
        Assert.Equal(4D, logicalGroup.ActualTextAnchorX);
        Assert.Equal(15D, logicalGroup.ActualTextAnchorY);

        byte[] pdf = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Drawing(clone)
            .ToBytes();
        string raw = Encoding.ASCII.GetString(pdf);

        Assert.Contains("/ActualText", raw, StringComparison.Ordinal);
        Assert.Contains("Vector label", PdfReadDocument.Open(pdf).ExtractText(), StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(OfficeFillRule.EvenOdd, " W* n")]
    [InlineData(OfficeFillRule.NonZero, " W n")]
    public void DrawingFlowPreservesFreeformGroupClipFillRule(OfficeFillRule fillRule, string expectedOperator) {
        var child = new OfficeDrawing(80D, 40D)
            .AddShape(new OfficeShape {
                Kind = OfficeShapeKind.Rectangle,
                Width = 80D,
                Height = 40D,
                FillColor = OfficeColor.FromRgb(37, 99, 235)
            }, 0D, 0D);
        OfficeClipPath clip = OfficeClipPath.Path(new[] {
            OfficePathCommand.MoveTo(0D, 0D),
            OfficePathCommand.LineTo(60D, 0D),
            OfficePathCommand.LineTo(60D, 40D),
            OfficePathCommand.LineTo(0D, 40D),
            OfficePathCommand.Close(),
            OfficePathCommand.MoveTo(15D, 10D),
            OfficePathCommand.LineTo(45D, 10D),
            OfficePathCommand.LineTo(45D, 30D),
            OfficePathCommand.LineTo(15D, 30D),
            OfficePathCommand.Close()
        }, fillRule);
        var drawing = new OfficeDrawing(80D, 50D).AddClippedDrawing(child, 10D, 5D, clip);

        byte[] pdf = PdfDocument.Create(new PdfOptions { CompressContentStreams = false })
            .Drawing(drawing)
            .ToBytes();

        Assert.Contains(expectedOperator, Encoding.ASCII.GetString(pdf), StringComparison.Ordinal);
    }
}
