using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public partial class DrawingTests {
    [Fact]
    public void OfficeDrawingSvgExporter_PreservesLegacyPointsAndOffersExplicitPixelSurface() {
        var drawing = new OfficeDrawing(120D, 80D);

        string points = OfficeDrawingSvgExporter.ToSvg(drawing, 1.5D);
        string pixels = OfficeDrawingSvgExporter.ToSvg(drawing, 1.5D, OfficeSvgSizeUnit.Pixel);

        Assert.Contains("width=\"180pt\" height=\"120pt\" viewBox=\"0 0 120 80\"", points, StringComparison.Ordinal);
        Assert.Contains("width=\"180px\" height=\"120px\" viewBox=\"0 0 120 80\"", pixels, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingText_RetainsResolvedAdvanceAndOverflowAcrossCloneAndSvg() {
        var drawing = new OfficeDrawing(120D, 40D).AddPositionedText(
            "One model.",
            4D,
            4D,
            54D,
            20D,
            new OfficeFontInfo("Arial", 14D),
            OfficeColor.Black);

        OfficeDrawing clone = drawing.Clone();
        OfficeDrawingText text = Assert.Single(clone.Elements.OfType<OfficeDrawingText>());
        string svg = OfficeDrawingSvgExporter.ToSvg(clone, 1D, OfficeSvgSizeUnit.Pixel);

        Assert.Equal(OfficeTextOverflowBehavior.Clip, text.OverflowBehavior);
        Assert.Equal(54D, text.TextAdvanceWidth);
        Assert.Contains("textLength=\"54\" lengthAdjust=\"spacingAndGlyphs\"", svg, StringComparison.Ordinal);
        Assert.Contains(">One model.</text>", svg, StringComparison.Ordinal);
    }

    [Fact]
    public void OfficeDrawingSvgExporter_PrefixesGeneratedResourceIdentifiersAndReferences() {
        var drawing = new OfficeDrawing(120D, 80D);
        var shape = OfficeShape.Rectangle(80D, 40D);
        shape.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.Red, OfficeColor.Blue);
        shape.ClipPath = OfficeClipPath.Rectangle(60D, 30D);
        drawing.AddShape(shape, 10D, 10D);

        string svg = OfficeDrawingSvgExporter.ToSvg(
            drawing,
            1D,
            OfficeSvgSizeUnit.Pixel,
            imageCodec: null,
            resourceIdPrefix: "page-2-");

        Assert.Contains("id=\"page-2-officeimo-gradient-1\"", svg, StringComparison.Ordinal);
        Assert.Contains("url(#page-2-officeimo-gradient-1)", svg, StringComparison.Ordinal);
        Assert.Contains("id=\"page-2-officeimo-clip-1\"", svg, StringComparison.Ordinal);
        Assert.Contains("url(#page-2-officeimo-clip-1)", svg, StringComparison.Ordinal);
        Assert.DoesNotContain("id=\"officeimo-", svg, StringComparison.Ordinal);
    }
}
