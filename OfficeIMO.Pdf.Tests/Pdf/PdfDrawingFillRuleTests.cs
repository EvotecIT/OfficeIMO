using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfDrawingFillRuleTests {
    [Theory]
    [InlineData(OfficeFillRule.EvenOdd, false, false)]
    [InlineData(OfficeFillRule.NonZero, false, false)]
    [InlineData(OfficeFillRule.EvenOdd, true, false)]
    [InlineData(OfficeFillRule.NonZero, true, false)]
    [InlineData(OfficeFillRule.EvenOdd, false, true)]
    [InlineData(OfficeFillRule.NonZero, false, true)]
    [InlineData(OfficeFillRule.EvenOdd, true, true)]
    [InlineData(OfficeFillRule.NonZero, true, true)]
    public void CompoundPathPreservesFillRuleInSolidAndTransformedPaint(
        OfficeFillRule fillRule, bool transformed, bool stroked) {
        OfficeShape shape = CreateCompoundPath(fillRule, transformed);
        if (stroked) {
            shape.StrokeColor = OfficeColor.Black;
            shape.StrokeWidth = 0.5D;
        }
        PdfDocument document = PdfDocument.Create(CreateOptions());
        document.Content.Shape(shape);

        byte[] pdf = document.ToBytes();
        string operation = stroked ? " B" : " f";
        if (fillRule == OfficeFillRule.EvenOdd) operation += "*";
        Assert.Contains(operation + "\n", Encoding.ASCII.GetString(pdf), StringComparison.Ordinal);
        OfficeDrawing drawing = PdfReadDocument.Open(pdf).Pages[0].ToDrawing();
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(drawing, background: OfficeColor.White);
        OfficeColor center = raster.GetPixel(15, 10);
        Assert.Equal(fillRule == OfficeFillRule.EvenOdd ? OfficeColor.White : OfficeColor.Green, center);
        Assert.Equal(OfficeColor.Green, raster.GetPixel(1, 10));
    }

    [Theory]
    [InlineData(OfficeFillRule.EvenOdd, false)]
    [InlineData(OfficeFillRule.NonZero, false)]
    [InlineData(OfficeFillRule.EvenOdd, true)]
    [InlineData(OfficeFillRule.NonZero, true)]
    public void CompoundGradientUsesShapeFillRuleForItsClip(OfficeFillRule fillRule, bool transformed) {
        OfficeShape shape = CreateCompoundPath(fillRule, transformed);
        shape.FillGradient = OfficeLinearGradient.Horizontal(OfficeColor.Green, OfficeColor.Blue);
        PdfDocument document = PdfDocument.Create(CreateOptions());
        document.Content.Shape(shape);

        string content = Encoding.ASCII.GetString(document.ToBytes());
        Assert.Contains(fillRule == OfficeFillRule.EvenOdd ? " W* n" : " W n", content, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData(OfficeFillRule.EvenOdd)]
    [InlineData(OfficeFillRule.NonZero)]
    public void HeaderCompoundPathPreservesFillRule(OfficeFillRule fillRule) {
        PdfDocument document = PdfDocument.Create(CreateOptions());
        document.Compose(builder => builder.Page(page => page
            .Header(header => header.Shape(CreateCompoundPath(fillRule, transformed: false)))
            .Content(content => content.Paragraph(paragraph => paragraph.Text("Body")))));

        string content = Encoding.ASCII.GetString(document.ToBytes(options:new global::OfficeIMO.Pdf.PdfSaveOptions { AllowLoss = true }));
        Assert.Contains(fillRule == OfficeFillRule.EvenOdd ? " f*\n" : " f\n", content, StringComparison.Ordinal);
    }

    [Theory]
    [InlineData("", false)]
    [InlineData("fill-rule='evenodd'", true)]
    public void SvgProducerDefaultsToNonzeroWhileExplicitEvenoddRetainsItsHole(string rule, bool hole) {
        byte[] svg = Encoding.UTF8.GetBytes("<svg xmlns='http://www.w3.org/2000/svg' width='40' height='30'><path fill='green' " + rule +
            " d='M0 0H30V20H0Z M3 0H33V20H3Z'/></svg>");
        Assert.True(OfficeSvgDrawingReader.TryRead(svg, out OfficeDrawing? drawing, out int unsupported));
        Assert.Equal(0, unsupported);
        byte[] pdf = PdfDocument.Create(CreateOptions()).Drawing(drawing!).ToBytes();
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfReadDocument.Open(pdf).Pages[0].ToDrawing(), background: OfficeColor.White);
        Assert.Equal(hole ? OfficeColor.White : OfficeColor.Green, raster.GetPixel(15, 10));
    }

    private static PdfCreateOptions CreateOptions() => new PdfCreateOptions {
        PageWidth = 80D, PageHeight = 80D,
        MarginLeft = 0D, MarginRight = 0D, MarginTop = 0D, MarginBottom = 0D,
        CompressContentStreams = false
    };

    private static OfficeShape CreateCompoundPath(OfficeFillRule fillRule, bool transformed) {
        OfficeShape shape = OfficeShape.Path(
            OfficePathCommand.MoveTo(0D, 0D), OfficePathCommand.LineTo(30D, 0D),
            OfficePathCommand.LineTo(30D, 20D), OfficePathCommand.LineTo(0D, 20D), OfficePathCommand.Close(),
            OfficePathCommand.MoveTo(3D, 0D), OfficePathCommand.LineTo(33D, 0D),
            OfficePathCommand.LineTo(33D, 20D), OfficePathCommand.LineTo(3D, 20D), OfficePathCommand.Close());
        shape.FillColor = OfficeColor.Green;
        shape.FillRule = fillRule;
        shape.StrokeWidth = 0D;
        if (transformed) shape.Transform = OfficeTransform.Identity;
        return shape;
    }
}
