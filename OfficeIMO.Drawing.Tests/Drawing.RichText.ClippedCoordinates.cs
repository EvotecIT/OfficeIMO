using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRichTextClippedCoordinatesTests {
    [Theory]
    [InlineData(-10, 0)]
    [InlineData(0, -10)]
    [InlineData(-10, -10)]
    public void Clipped_rich_text_keeps_bleed_coordinates_through_master_projection_copy_and_tint(double x, double y) {
        var paragraphs = new[] { new OfficeRichTextParagraph(new[] {
            new OfficeRichTextRun("Bleeding text", 12, OfficeColor.Blue, bold: true)
        }, OfficeTextAlignment.Center) };
        var local = new OfficeDrawing(60, 40).AddRichTextParagraphs(paragraphs, 0, 0, 60, 40);
        var scene = new OfficeDrawing(100, 100).AddDrawingForClippedRendering(local, x, y, null);
        var masterProjection = new OfficeDrawing(100, 100).AddDrawingForClippedRendering(scene, 0, 0, null);
        var clipped = new OfficeDrawing(100, 100).AddClippedDrawing(masterProjection, 0, 0, OfficeClipPath.Rectangle(100, 100));
        var tinted = clipped.Clone();
        tinted.ApplyColorTint(OfficeColor.Red);
        foreach (OfficeDrawing drawing in new[] { scene, scene.Clone(), masterProjection, clipped.Clone(), tinted }) {
            OfficeDrawingRichText frame = Assert.Single(Frames(drawing));
            Assert.Equal(x, frame.X);
            Assert.Equal(y, frame.Y);
            Assert.Equal("Bleeding text", frame.PlainText);
            Assert.True(frame.Runs[0].Bold);
            Assert.Equal(OfficeTextAlignment.Center, frame.Paragraphs[0].Alignment);
        }
        Assert.Contains("clip-path=", OfficeDrawingSvgExporter.ToSvg(clipped));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeDrawing(100, 100).AddDrawing(scene, 0, 0));
    }

    [Fact]
    public void Authored_rich_text_and_clipped_intermediates_retain_their_coordinate_validation() {
        var runs = new[] { new OfficeRichTextRun("Text", 12, OfficeColor.Black) };
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeDrawingRichText(runs, -1, 0, 40, 20));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeDrawingRichText(runs, 0, -1, 40, 20));
        var local = new OfficeDrawing(40, 20).AddRichText(runs, 0, 0, 40, 20);
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeDrawing(100, 100).AddDrawingForClippedRendering(local, double.NaN, 0, null));
        Assert.Throws<ArgumentOutOfRangeException>(() => new OfficeDrawing(100, 100).AddDrawingForClippedRendering(local, 0, double.NegativeInfinity, null));
    }

    private static IEnumerable<OfficeDrawingRichText> Frames(OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            if (element is OfficeDrawingRichText text) yield return text;
            if (element is OfficeDrawingGroup group) foreach (OfficeDrawingRichText child in Frames(group.Drawing)) yield return child;
        }
    }
}
