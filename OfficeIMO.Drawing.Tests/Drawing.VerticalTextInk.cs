using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingVerticalTextInkTests {
    [Theory]
    [InlineData(false, false, false)]
    [InlineData(false, false, true)]
    [InlineData(false, true, false)]
    [InlineData(false, true, true)]
    [InlineData(true, false, false)]
    [InlineData(true, false, true)]
    [InlineData(true, true, false)]
    [InlineData(true, true, true)]
    public void VerticalInkUsesShapedContoursAndActualTextBoxClip(bool rotated, bool color, bool clipped) {
        byte[] font = color ? ManagedTextShapingTestAssets.CreateColorFont('A') : ManagedTextShapingTestAssets.CreateFont('A');
        var source = new OfficeDrawing(100, 100).AddFont("Vertical Ink", font);
        source.AddVerticalText("AA", 20, 10, 50, clipped ? 25 : 75,
            new OfficeFontInfo("Vertical Ink", 30, OfficeFontStyle.Bold | OfficeFontStyle.Italic | OfficeFontStyle.Underline));
        var drawing = rotated ? new OfficeDrawing(100, 100).AddDrawing(source, 0, 0, new OfficeImageFrameTransform(27, 50, 50)) : source;
        var provider = new VerticalProvider();
        var bounds = new List<(double Left, double Top, double Right, double Bottom)>();
        bool observedClip = false;
        new OfficeRasterCanvas(new OfficeRasterImage(1, 1), font: null, fonts: null, textShapingProvider: provider).InspectDrawingTextInk(
            drawing, OfficeTransform.Identity, Array.Empty<OfficeTextInkClip>(), (ink, reason) => {
                Assert.True(ink.IsMeasured, reason); observedClip |= ink.IsClipped;
                if (ink.HasInk) bounds.Add((ink.Left, ink.Top, ink.Right, ink.Bottom));
            });
        Assert.NotEmpty(bounds);
        // The conservative underline envelope can itself touch the top clip.
        if (clipped) Assert.True(observedClip);
        var image = OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions { TextShapingProvider = provider });
        var pixels = Enumerable.Range(0, 100).SelectMany(y => Enumerable.Range(0, 100).Select(x => (X: x, Y: y)))
            .Where(p => image.GetPixel(p.X, p.Y).A != 0).ToArray();
        Assert.NotEmpty(pixels);
        // Rotated text is sampled from an intermediate image; nominal geometry
        // and interpolated pixel edges need not be identical.
        Assert.InRange(Math.Abs(pixels.Min(p => p.X) - bounds.Min(b => b.Left)), 0, 3);
        Assert.InRange(Math.Abs(pixels.Min(p => p.Y) - bounds.Min(b => b.Top)), 0, 3);
        Assert.InRange(Math.Abs(pixels.Max(p => p.X) + 1 - bounds.Max(b => b.Right)), 0, 3);
        Assert.InRange(Math.Abs(pixels.Max(p => p.Y) + 1 - bounds.Max(b => b.Bottom)), 0, 3);
    }

    [Fact]
    public void VerticalInputIsBoundedBeforeShaping() {
        var drawing = new OfficeDrawing(100, 100).AddFont("Vertical Ink", ManagedTextShapingTestAssets.CreateFont('A'));
        drawing.AddVerticalText(new string('A', 65537), 0, 0, 100, 100, new OfficeFontInfo("Vertical Ink", 30));
        var provider = new VerticalProvider();
        var canvas = new OfficeRasterCanvas(new OfficeRasterImage(1, 1), font: null, fonts: null, textShapingProvider: provider);
        Assert.Throws<NotSupportedException>(() => canvas.InspectDrawingTextInk(drawing, OfficeTransform.Identity,
            Array.Empty<OfficeTextInkClip>(), (_, _) => { }));
        Assert.Equal(0, provider.Calls);
    }

    private sealed class VerticalProvider : IOfficeTextShapingProvider {
        public int Calls { get; private set; }
        public OfficeTextShapingResult? ShapeText(OfficeTextShapingRequest request) {
            Calls++;
            return new OfficeTextShapingResult(new[] {
            new OfficeShapedGlyph(1, "A", 0, advanceWidth: 0, advanceHeight: -1000, offsetX: 0, offsetY: 0),
            new OfficeShapedGlyph(1, "A", 1, advanceWidth: 0, advanceHeight: -1000, offsetX: 0, offsetY: 0)
            }, OfficeTextDirection.TopToBottom);
        }
    }
}
