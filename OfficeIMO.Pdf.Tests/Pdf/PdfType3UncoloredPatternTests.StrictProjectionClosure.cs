using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public partial class PdfType3UncoloredPatternTests {
    [Fact]
    public void RenderPage_FailsClosedForMalformedOrdinaryType3ImageDecodeArray() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            type3PaintType: 1,
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8 /Decode [0 1 1 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void CurvedPathIntersectionCannotRemainExactAfterFlattening() {
        OfficePathCommand[] curved = {
            OfficePathCommand.MoveTo(0D, 0D),
            OfficePathCommand.CubicBezierTo(0D, 10D, 10D, 10D, 10D, 0D),
            OfficePathCommand.LineTo(0D, 0D),
            OfficePathCommand.Close()
        };
        OfficePathCommand[] triangle = {
            OfficePathCommand.MoveTo(0D, 0D),
            OfficePathCommand.LineTo(10D, 0D),
            OfficePathCommand.LineTo(5D, 10D),
            OfficePathCommand.Close()
        };
        Assert.True(PdfPageClipPath.TryCreatePath(curved, OfficeFillRule.NonZero, out PdfPageClipPath first));
        Assert.True(PdfPageClipPath.TryCreatePath(triangle, OfficeFillRule.NonZero, out PdfPageClipPath second));

        PdfPageClipPath intersection = PdfPageClipPath.ResolveActiveClip(first, second);

        Assert.False(intersection.IsExact);
    }

    [Fact]
    public void VisualParser_RejectsAxialShadingCollapsedByRendererTolerance() {
        var shading = new PdfPageShadingResource(
            0D, 0D, 0.05D, 0D,
            OfficeColor.Red,
            OfficeColor.Blue);

        bool supported = PdfPageContentVisualParser.IsSupportedExactShadingPlacement(
            shading,
            Matrix2D.Identity,
            0D,
            0D,
            100D,
            100D,
            100D);

        Assert.False(supported);
    }

    [Fact]
    public void RenderPage_CropsSubToleranceType3ImageAtPageBoundary() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "BT /FType3 18 Tf 0 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 q 500 0 0 700 -0.05 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            type3PaintType: 1,
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        OfficeDrawingImage image = Assert.Single(drawing.Images);
        Assert.True(image.Projection.HasCrop, $"x={image.Projection.X:R}; width={image.Projection.Width:R}; sourceLeft={image.Projection.SourceLeft:R}; sourceWidth={image.Projection.SourceWidth:R}");
        Assert.True(image.Projection.SourceLeft > 0D);
    }
}
