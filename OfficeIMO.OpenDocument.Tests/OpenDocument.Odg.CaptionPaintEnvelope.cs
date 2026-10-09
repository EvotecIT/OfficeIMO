using System.Linq;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.OpenDocument.Tests;

public sealed partial class OpenDocumentOdgLineLabelTests {
    [Theory]
    [InlineData("left", "f")]
    [InlineData("center", "f")]
    [InlineData("right", "f")]
    [InlineData("justify", "f")]
    [InlineData("left", "j")]
    [InlineData("center", "j")]
    [InlineData("right", "j")]
    [InlineData("justify", "j")]
    public void CaptionPaintCanvasRetainsItalicOverhangWithoutMovingItsLogicalFrame(string area, string text) {
        foreach (string font in new[] { "Arial", "Times New Roman" }) {
            var document = AreaCaption(area, "center", true, 1);
            var shape = document.Pages[0].Shapes[0];
            shape.Text = text; var paragraph = Assert.Single(shape.Paragraphs);
            paragraph.FontFamily = font; paragraph.FontSize = OdfLength.Points(30);
            paragraph.Italic = true; paragraph.TextAlign = "center";
            foreach (var read in RoundTrips(document)) {
                var drawing = read.Pages[0].ToDrawing().Value;
                var group = Assert.Single(drawing.Elements.OfType<OfficeDrawingEffectGroup>());
                var frame = Assert.Single(group.Drawing.Elements.OfType<OfficeDrawingRichText>());
                double advance = new OfficeRasterCanvas(new OfficeRasterImage(1, 1)).MeasureText(text, 30, font, OfficeFontStyle.Italic);
                Assert.Equal(advance, frame.Width, 6);
                double x = area switch { "left" => 0, "right" => 1 - advance, _ => (1 - advance) / 2 };
                var expected = OfficeTransform.RotateDegrees(System.Math.Atan2(.8, .6) * 180 / System.Math.PI)
                    .Then(OfficeTransform.Translate(100, 100)).TransformPoint(new OfficePoint(x, 0));
                var actual = group.Transform.TransformPoint(new OfficePoint(frame.X, frame.Y + frame.Height / 2));
                Assert.Equal(expected.X, actual.X, 6); Assert.Equal(expected.Y, actual.Y, 6);

                // A larger canvas with the identical logical frame provides real glyph-ink
                // evidence. The intermediate must contain every reference pixel at both scales.
                var reference = new OfficeDrawing(group.Drawing.Width + 80, group.Drawing.Height);
                reference.AddRichTextParagraphs(frame.Paragraphs, frame.X + 40, frame.Y, frame.Width, frame.Height,
                    frame.VerticalAlignment, wrapText: false, padding: frame.Padding);
                foreach (int scale in new[] { 1, 4 }) {
                    var image = OfficeDrawingRasterRenderer.Render(group.Drawing, scale);
                    var expanded = OfficeDrawingRasterRenderer.Render(reference, scale);
                    int ink = 0;
                    for (int y = 0; y < expanded.Height; y++) for (int column = 0; column < expanded.Width; column++) {
                        OfficeColor pixel = expanded.GetPixel(column, y);
                        if (pixel.A > 0) ink++;
                        int local = column - 40 * scale;
                        if (local >= 0 && local < image.Width) Assert.Equal(pixel, image.GetPixel(local, y));
                        else Assert.Equal(0, pixel.A);
                    }
                    Assert.True(ink > 0);
                }
            }
        }
    }
}
