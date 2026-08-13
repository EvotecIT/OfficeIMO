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
    public void ConvexClipIntersectionDoesNotRetainSubToleranceOutsideVertices() {
        OfficePathCommand[] subject = {
            OfficePathCommand.MoveTo(-0.0001D, 1D),
            OfficePathCommand.LineTo(5D, 1D),
            OfficePathCommand.LineTo(5D, 5D),
            OfficePathCommand.Close()
        };
        OfficePathCommand[] clip = {
            OfficePathCommand.MoveTo(0D, 0D),
            OfficePathCommand.LineTo(10D, 0D),
            OfficePathCommand.LineTo(0D, 10D),
            OfficePathCommand.Close()
        };
        Assert.True(PdfPageClipPath.TryCreatePath(subject, OfficeFillRule.NonZero, out PdfPageClipPath first));
        Assert.True(PdfPageClipPath.TryCreatePath(clip, OfficeFillRule.NonZero, out PdfPageClipPath second));

        PdfPageClipPath intersection = PdfPageClipPath.ResolveActiveClip(first, second);

        Assert.True(intersection.IsExact);
        Assert.All(
            intersection.Commands.Where(command => command.Kind != OfficePathCommandKind.Close),
            command => Assert.True(command.Point.X >= 0D, $"Unexpected outside x-coordinate {command.Point.X:R}."));
    }

    [Fact]
    public void ShallowConcaveClipIntersectionCannotRemainExact() {
        OfficePathCommand[] firstCommands = {
            OfficePathCommand.MoveTo(0D, 0D),
            OfficePathCommand.LineTo(10D, 0D),
            OfficePathCommand.LineTo(10D, 10D),
            OfficePathCommand.LineTo(5D, 9.99995D),
            OfficePathCommand.LineTo(0D, 10D),
            OfficePathCommand.Close()
        };
        OfficePathCommand[] secondCommands = {
            OfficePathCommand.MoveTo(2D, 2D),
            OfficePathCommand.LineTo(12D, 2D),
            OfficePathCommand.LineTo(12D, 12D),
            OfficePathCommand.LineTo(7D, 11.99995D),
            OfficePathCommand.LineTo(2D, 12D),
            OfficePathCommand.Close()
        };
        Assert.True(PdfPageClipPath.TryCreatePath(firstCommands, OfficeFillRule.NonZero, out PdfPageClipPath first));
        Assert.True(PdfPageClipPath.TryCreatePath(secondCommands, OfficeFillRule.NonZero, out PdfPageClipPath second));

        PdfPageClipPath intersection = PdfPageClipPath.ResolveActiveClip(first, second);

        Assert.False(intersection.IsExact);
    }

    [Fact]
    public void ExactClipIntersectionDoesNotUseNearParallelEndpointFallback() {
        OfficePathCommand[] subjectCommands = {
            OfficePathCommand.MoveTo(1D, -0.00000001D),
            OfficePathCommand.LineTo(2D, 0.00000001D),
            OfficePathCommand.LineTo(3D, 2D),
            OfficePathCommand.LineTo(1D, 2D),
            OfficePathCommand.Close()
        };
        OfficePathCommand[] clipCommands = {
            OfficePathCommand.MoveTo(0D, 0D),
            OfficePathCommand.LineTo(10D, 0D),
            OfficePathCommand.LineTo(10D, 10D),
            OfficePathCommand.LineTo(0D, 10D),
            OfficePathCommand.Close()
        };
        Assert.True(PdfPageClipPath.TryCreatePath(subjectCommands, OfficeFillRule.NonZero, out PdfPageClipPath subject));
        Assert.True(PdfPageClipPath.TryCreatePath(clipCommands, OfficeFillRule.NonZero, out PdfPageClipPath clip));

        PdfPageClipPath intersection = PdfPageClipPath.ResolveActiveClip(subject, clip);

        Assert.True(intersection.IsExact);
        Assert.Contains(
            intersection.Commands,
            command => command.Kind != OfficePathCommandKind.Close && command.Point.Y == 0D);
    }

    [Fact]
    public void ExactClipIntersectionPreservesSkinnyClosingVertex() {
        OfficePathCommand[] subjectCommands = {
            OfficePathCommand.MoveTo(0D, 0D),
            OfficePathCommand.LineTo(10D, 0D),
            OfficePathCommand.LineTo(0.0005D, 0.0005D),
            OfficePathCommand.Close()
        };
        OfficePathCommand[] clipCommands = {
            OfficePathCommand.MoveTo(-1D, -1D),
            OfficePathCommand.LineTo(20D, -1D),
            OfficePathCommand.LineTo(20D, 20D),
            OfficePathCommand.LineTo(-1D, 20D),
            OfficePathCommand.Close()
        };
        Assert.True(PdfPageClipPath.TryCreatePath(subjectCommands, OfficeFillRule.NonZero, out PdfPageClipPath subject));
        Assert.True(PdfPageClipPath.TryCreatePath(clipCommands, OfficeFillRule.NonZero, out PdfPageClipPath clip));

        PdfPageClipPath intersection = PdfPageClipPath.ResolveActiveClip(subject, clip);

        Assert.True(intersection.IsExact);
        Assert.True(intersection.Width > 0D);
        Assert.True(intersection.Height > 0D);
        Assert.Contains(intersection.Commands, command => command.Kind == OfficePathCommandKind.LineTo && command.Point.X == 0.0005D);
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

    [Theory]
    [InlineData("/Width 1.5 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8")]
    [InlineData("/Width 1 /Height 1.5 /ColorSpace /DeviceGray /BitsPerComponent 8")]
    [InlineData("/Width 1 /Height 1 /ColorSpace /DeviceGray /BitsPerComponent 8.5")]
    public void RenderPage_FailsClosedForFractionalType3ImageDimensions(string imageEntries) {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 q 500 0 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            type3PaintType: 1,
            extraObjects: new[] { StreamObject(8, "<< /Type /XObject /Subtype /Image " + imageEntries, "x") });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_ClipsSubTolerancePatternedImageMaskExactly() {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 5 5] /XStep 5 /YStep 5 /Resources << >>",
            patternContent: "1 0 0 rg 0 0 5 5 re f",
            glyphContent: "500 0 d0 0.00001 0 499.99999 700 re W n q 500 0 0 700 0 0 cm /Im1 Do Q",
            glyphResources: "<< /XObject << /Im1 8 0 R >> >>",
            extraObjects: new[] {
                StreamObject(8, "<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Decode [1 0]", "x")
            });

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.DoesNotContain(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
        Assert.Single(drawing.Elements.OfType<OfficeDrawingEffectGroup>());
    }

    [Theory]
    [InlineData("/Domain [0.0000000005 1]")]
    [InlineData("/Domain [0 1] /Range [0.0000000005 1 0 1 0 1]")]
    public void RenderPage_FailsClosedForNearlyCanonicalType3ShadingIntervals(string intervalEntries) {
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function << /FunctionType 2 " + intervalEntries + " /C0 [1 0 0] /C1 [0 0 1] /N 1 >> /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }

    [Fact]
    public void RenderPage_FailsClosedForDiscontinuousStitchedType3ShadingBoundary() {
        const string function =
            "<< /FunctionType 3 /Domain [0 1] " +
            "/Functions [" +
            "<< /FunctionType 2 /Domain [0 1] /C0 [1 0 0] /C1 [0 1 0] /N 1 >> " +
            "<< /FunctionType 2 /Domain [0 1] /C0 [0 0 1] /C1 [1 1 1] /N 1 >>] " +
            "/Bounds [0.5] /Encode [0 1 0 1] >>";
        byte[] pdf = BuildUncoloredType3PatternPdf(
            pageContent: "/Pattern cs /P1 scn BT /FType3 18 Tf 20 100 Td (A) Tj ET",
            pageColorSpaceResources: string.Empty,
            patternDictionary: "<< /Type /Pattern /PatternType 2 /Shading << /ShadingType 2 /ColorSpace /DeviceRGB /Coords [20 100 30 100] /Function " + function + " /Extend [true true] >>",
            patternContent: string.Empty,
            patternIsStream: false);

        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(pdf));

        Assert.Contains(result.CapabilityDiagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.Type3FontSubstitutionId);
    }
}
