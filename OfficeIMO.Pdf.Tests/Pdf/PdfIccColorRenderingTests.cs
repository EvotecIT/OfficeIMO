using System.Globalization;
using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfIccColorRenderingTests {
    [Fact]
    public void RenderPage_AppliesEmbeddedCmykMabProfileWithoutApproximationDiagnostic() {
        byte[] pdf = BuildIccContentPdf(
            IccMabTestProfiles.CreateCmykLab8(),
            "/N 4",
            "0 1 1 0 scn");

        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeImageExportResult result = page.ExportImage(OfficeImageExportFormat.Png);

        OfficeColor fill = Assert.Single(drawing.Shapes).Shape.FillColor!.Value;
        Assert.InRange(fill.R, 245, 255);
        Assert.InRange(fill.G, 0, 15);
        Assert.InRange(fill.B, 0, 15);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.IccColorSpaceId);
    }

    [Fact]
    public void RenderPage_AppliesEmbeddedCmykLabLut8ProfileWithoutApproximationDiagnostic() {
        byte[] pdf = BuildIccContentPdf(
            IccLutTestProfiles.CreateCmykLut8(),
            "/N 4",
            "0 1 1 0 scn");

        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeImageExportResult result = page.ExportImage(OfficeImageExportFormat.Png);

        OfficeColor fill = Assert.Single(drawing.Shapes).Shape.FillColor!.Value;
        Assert.InRange(fill.R, 245, 255);
        Assert.InRange(fill.G, 0, 15);
        Assert.InRange(fill.B, 0, 15);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.IccColorSpaceId);
    }

    [Fact]
    public void RenderPage_AppliesContentAndExtGStateRenderingIntent() {
        byte[] profile = IccLutTestProfiles.CreateCmykLut8WithDistinctRelativeIntent();
        OfficeColor perceptual = Assert.Single(PdfPageImageRenderer.RenderPage(BuildIccContentPdf(
            profile,
            "/N 4",
            "/Perceptual ri 0 0 0 0 scn")).Shapes).Shape.FillColor!.Value;
        OfficeColor relative = Assert.Single(PdfPageImageRenderer.RenderPage(BuildIccContentPdf(
            profile,
            "/N 4",
            "0 0 0 0 scn")).Shapes).Shape.FillColor!.Value;
        OfficeColor extGStatePerceptual = Assert.Single(PdfPageImageRenderer.RenderPage(BuildIccContentPdf(
            profile,
            "/N 4",
            "/IntentGs gs 0 0 0 0 scn",
            extraResourceEntries: "/ExtGState << /IntentGs << /RI /Perceptual >> >>")).Shapes).Shape.FillColor!.Value;
        OfficeColor lateRelative = Assert.Single(PdfPageImageRenderer.RenderPage(BuildIccContentPdf(
            profile,
            "/N 4",
            "0 0 0 0 scn /RelativeColorimetric ri")).Shapes).Shape.FillColor!.Value;
        OfficeColor lateExtGStatePerceptual = Assert.Single(PdfPageImageRenderer.RenderPage(BuildIccContentPdf(
            profile,
            "/N 4",
            "/RelativeColorimetric ri 0 0 0 0 scn /IntentGs gs",
            extraResourceEntries: "/ExtGState << /IntentGs << /RI /Perceptual >> >>")).Shapes).Shape.FillColor!.Value;

        Assert.NotEqual(perceptual, relative);
        Assert.Equal(perceptual, extGStatePerceptual);
        Assert.Equal(relative, lateRelative);
        Assert.Equal(perceptual, lateExtGStatePerceptual);
    }

    [Fact]
    public void RenderPage_ResolvesIndirectExtGStateIntentAndTreatsNullAsAbsent() {
        byte[] profile = IccLutTestProfiles.CreateCmykLut8WithDistinctRelativeIntent();
        const string content =
            "/Perceptual ri /NullIntent gs /CsIcc cs 0 0 0 0 scn 0 0 10 10 re f\n" +
            "/RelativeColorimetric ri /IndirectIntent gs /CsIcc cs 0 0 0 0 scn 20 0 10 10 re f";
        OfficeColor[] colors = PdfPageImageRenderer.RenderPage(BuildIccContentPdf(
            profile,
            "/N 4",
            string.Empty,
            extraObjects: "7 0 obj\n8 0 R\nendobj\n8 0 obj\n/Perceptual\nendobj\n",
            extraResourceEntries:
                "/ExtGState << /NullIntent << /RI null >> /IndirectIntent << /RI 7 0 R >> >>",
            contentOverride: content)).Shapes.Select(shape => shape.Shape.FillColor!.Value).ToArray();

        Assert.Equal(2, colors.Length);
        Assert.Equal(colors[0], colors[1]);
    }

    [Fact]
    public void RenderPage_RestoresRenderingIntentAcrossGraphicsStateAndPropagatesIntoForm() {
        byte[] profile = IccLutTestProfiles.CreateCmykLut8WithDistinctRelativeIntent();
        const string twoShapes =
            "q /Perceptual ri /CsIcc cs 0 0 0 0 scn 0 0 10 10 re f Q\n" +
            "/CsIcc cs 0 0 0 0 scn 20 0 10 10 re f";
        OfficeDrawing stateDrawing = PdfPageImageRenderer.RenderPage(BuildIccContentPdf(
            profile,
            "/N 4",
            string.Empty,
            contentOverride: twoShapes));
        OfficeColor[] stateColors = stateDrawing.Shapes.Select(shape => shape.Shape.FillColor!.Value).ToArray();

        const string formContent = "/CsIcc cs 0 0 0 0 scn 0 0 10 10 re f";
        string formObject =
            "7 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Length " +
            Encoding.ASCII.GetByteCount(formContent).ToString(CultureInfo.InvariantCulture) +
            " >>\nstream\n" + formContent + "\nendstream\nendobj\n";
        OfficeColor formPerceptual = Assert.Single(PdfPageImageRenderer.RenderPage(BuildIccContentPdf(
            profile,
            "/N 4",
            string.Empty,
            extraObjects: formObject,
            extraResourceEntries: "/XObject << /Fm 7 0 R >>",
            contentOverride: "/Perceptual ri /Fm Do")).Shapes).Shape.FillColor!.Value;
        const string inheritedFormContent = "/RelativeColorimetric ri 0 0 10 10 re f";
        string inheritedFormObject =
            "7 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 10 10] /Length " +
            Encoding.ASCII.GetByteCount(inheritedFormContent).ToString(CultureInfo.InvariantCulture) +
            " >>\nstream\n" + inheritedFormContent + "\nendstream\nendobj\n";
        OfficeColor inheritedFormRelative = Assert.Single(PdfPageImageRenderer.RenderPage(BuildIccContentPdf(
            profile,
            "/N 4",
            string.Empty,
            extraObjects: inheritedFormObject,
            extraResourceEntries: "/XObject << /Fm 7 0 R >>",
            contentOverride: "/Perceptual ri /CsIcc cs 0 0 0 0 scn /Fm Do")).Shapes).Shape.FillColor!.Value;

        Assert.Equal(2, stateColors.Length);
        Assert.NotEqual(stateColors[0], stateColors[1]);
        Assert.Equal(stateColors[0], formPerceptual);
        Assert.Equal(stateColors[1], inheritedFormRelative);
    }

    [Fact]
    public void GetTextSpans_AppliesRenderingIntentToIccTextPaint() {
        byte[] profile = IccLutTestProfiles.CreateCmykLut8WithDistinctRelativeIntent();
        const string text = "BT /F1 12 Tf /CsIcc cs 0 0 0 0 scn /Perceptual ri 40 80 Td (X) Tj ET";
        byte[] pdf = BuildIccContentPdf(
            profile,
            "/N 4",
            string.Empty,
            extraResourceEntries: "/Font << /F1 << /Type /Font /Subtype /Type1 /BaseFont /Helvetica >> >>",
            contentOverride: text);

        PdfTextSpan span = Assert.Single(PdfReadDocument.Open(pdf).Pages[0].GetTextSpans());
        OfficeColor expected = Assert.Single(PdfPageImageRenderer.RenderPage(BuildIccContentPdf(
            profile,
            "/N 4",
            "/Perceptual ri 0 0 0 0 scn")).Shapes).Shape.FillColor!.Value;

        Assert.Equal(expected, span.Color);
    }

    [Fact]
    public void RenderPage_AppliesPaintTimeIntentToInheritedImageMaskTint() {
        byte[] profile = IccLutTestProfiles.CreateCmykLut8WithDistinctRelativeIntent();
        OfficeColor perceptual = ReadSingleRenderedImagePixel(BuildIccImageMaskPdf(
            profile,
            "/N 4",
            "/CsIcc cs 0 0 0 0 scn /RelativeColorimetric ri /Perceptual ri q 40 0 0 40 40 80 cm /Im1 Do Q"));
        OfficeColor relative = ReadSingleRenderedImagePixel(BuildIccImageMaskPdf(
            profile,
            "/N 4",
            "/CsIcc cs 0 0 0 0 scn /Perceptual ri /RelativeColorimetric ri q 40 0 0 40 40 80 cm /Im1 Do Q"));

        Assert.NotEqual(perceptual, relative);
    }

    [Fact]
    public void RenderPage_AppliesRenderingIntentToIccShadingStops() {
        byte[] profile = IccLutTestProfiles.CreateCmykLut8WithDistinctRelativeIntent();
        const string shadingObjects =
            "7 0 obj\n<< /ShadingType 2 /ColorSpace [/ICCBased 5 0 R] /Coords [0 0 100 0] /Function 8 0 R >>\nendobj\n" +
            "8 0 obj\n<< /FunctionType 2 /Domain [0 1] /C0 [0 0 0 0] /C1 [1 1 1 1] /N 1 >>\nendobj\n";
        OfficeDrawing perceptualDrawing = PdfPageImageRenderer.RenderPage(BuildIccContentPdf(
            profile,
            "/N 4",
            string.Empty,
            extraObjects: shadingObjects,
            extraResourceEntries: "/Shading << /Sh 7 0 R >>",
            contentOverride: "/Perceptual ri /Sh sh"));
        OfficeDrawing relativeDrawing = PdfPageImageRenderer.RenderPage(BuildIccContentPdf(
            profile,
            "/N 4",
            string.Empty,
            extraObjects: shadingObjects,
            extraResourceEntries: "/Shading << /Sh 7 0 R >>",
            contentOverride: "/RelativeColorimetric ri /Sh sh"));

        OfficeColor perceptual = Assert.Single(perceptualDrawing.Shapes).Shape.FillGradient!.Stops[0].Color;
        OfficeColor relative = Assert.Single(relativeDrawing.Shapes).Shape.FillGradient!.Stops[0].Color;
        Assert.NotEqual(perceptual, relative);
    }

    [Fact]
    public void RenderPage_AppliesRenderingIntentToIccShadingPatternStops() {
        byte[] profile = IccLutTestProfiles.CreateCmykLut8WithDistinctRelativeIntent();
        const string shadingObjects =
            "7 0 obj\n<< /Type /Pattern /PatternType 2 /Shading 8 0 R >>\nendobj\n" +
            "8 0 obj\n<< /ShadingType 2 /ColorSpace [/ICCBased 5 0 R] /Coords [0 0 100 0] /Function 9 0 R >>\nendobj\n" +
            "9 0 obj\n<< /FunctionType 2 /Domain [0 1] /C0 [0 0 0 0] /C1 [1 1 1 1] /N 1 >>\nendobj\n";
        OfficeDrawing perceptualDrawing = PdfPageImageRenderer.RenderPage(BuildIccContentPdf(
            profile,
            "/N 4",
            string.Empty,
            extraObjects: shadingObjects,
            extraResourceEntries: "/Pattern << /Sp 7 0 R >>",
            contentOverride: "/RelativeColorimetric ri /Pattern cs /Sp scn /Perceptual ri 0 0 100 100 re f"));
        OfficeDrawing relativeDrawing = PdfPageImageRenderer.RenderPage(BuildIccContentPdf(
            profile,
            "/N 4",
            string.Empty,
            extraObjects: shadingObjects,
            extraResourceEntries: "/Pattern << /Sp 7 0 R >>",
            contentOverride: "/Perceptual ri /Pattern cs /Sp scn /RelativeColorimetric ri 0 0 100 100 re f"));

        OfficeColor perceptual = Assert.Single(perceptualDrawing.Shapes).Shape.FillGradient!.Stops[0].Color;
        OfficeColor relative = Assert.Single(relativeDrawing.Shapes).Shape.FillGradient!.Stops[0].Color;
        Assert.NotEqual(perceptual, relative);
    }

    [Fact]
    public void RenderPage_AppliesRenderingIntentInsideTilingPattern() {
        byte[] profile = IccLutTestProfiles.CreateCmykLut8WithDistinctRelativeIntent();
        const string tileContent = "/CsIcc cs 0 0 0 0 scn 0 0 10 10 re f";
        string patternObject =
            "7 0 obj\n<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 10 10] " +
            "/XStep 10 /YStep 10 /Resources << /ColorSpace << /CsIcc [/ICCBased 5 0 R] >> >> /Length " +
            Encoding.ASCII.GetByteCount(tileContent).ToString(CultureInfo.InvariantCulture) +
            " >>\nstream\n" + tileContent + "\nendstream\nendobj\n";
        OfficeRasterImage perceptual = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(BuildIccContentPdf(
            profile,
            "/N 4",
            string.Empty,
            extraObjects: patternObject,
            extraResourceEntries: "/Pattern << /P 7 0 R >>",
            contentOverride: "/RelativeColorimetric ri /Pattern cs /P scn /Perceptual ri 0 0 100 100 re f")));
        OfficeRasterImage relative = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(BuildIccContentPdf(
            profile,
            "/N 4",
            string.Empty,
            extraObjects: patternObject,
            extraResourceEntries: "/Pattern << /P 7 0 R >>",
            contentOverride: "/Perceptual ri /Pattern cs /P scn /RelativeColorimetric ri 0 0 100 100 re f")));

        Assert.NotEqual(perceptual.GetPixel(50, 150), relative.GetPixel(50, 150));
    }

    [Fact]
    public void RenderPage_RendersSharedSoftMaskUnderEachInheritedRenderingIntent() {
        byte[] profile = IccLutTestProfiles.CreateCmykLut8WithDistinctRelativeIntent();
        const string maskContent = "/CsIcc cs 0 0 0 0 scn 0 0 240 200 re f";
        string maskObjects =
            "7 0 obj\n<< /Type /XObject /Subtype /Form /BBox [0 0 240 200] /Group << /S /Transparency /CS /DeviceRGB >> " +
            "/Resources << /ColorSpace << /CsIcc [/ICCBased 5 0 R] >> >> /Length " +
            Encoding.ASCII.GetByteCount(maskContent).ToString(CultureInfo.InvariantCulture) +
            " >>\nstream\n" + maskContent + "\nendstream\nendobj\n" +
            "8 0 obj\n<< /S /Luminosity /G 7 0 R >>\nendobj\n";
        const string content =
            "q /Perceptual ri /Mask gs 1 0 0 rg 0 0 100 100 re f Q\n" +
            "q /RelativeColorimetric ri /Mask gs 1 0 0 rg 120 0 100 100 re f Q";
        OfficeRasterImage raster = OfficeDrawingRasterRenderer.Render(PdfPageImageRenderer.RenderPage(BuildIccContentPdf(
            profile,
            "/N 4",
            string.Empty,
            extraObjects: maskObjects,
            extraResourceEntries: "/ExtGState << /Mask << /SMask 8 0 R >> >>",
            contentOverride: content)));

        Assert.NotEqual(raster.GetPixel(50, 150), raster.GetPixel(170, 150));
    }

    [Fact]
    public void RenderPage_AppliesEmbeddedMatrixTrcProfileWithoutApproximationDiagnostic() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        SwapTagPayload(profile, "rXYZ", "bXYZ");
        byte[] pdf = BuildIccContentPdf(profile, "/N 3 /Range [0 1 0 1 0 1]", "1 0 0 scn");

        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeImageExportResult result = page.ExportImage(OfficeImageExportFormat.Png);

        OfficeColor fill = Assert.Single(drawing.Shapes).Shape.FillColor!.Value;
        Assert.True(fill.B > 240, "The swapped ICC matrix should map the red device channel to blue.");
        Assert.True(fill.R < 40, "The embedded ICC matrix must replace the declared-component RGB fallback.");
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.IccColorSpaceId);
    }

    [Fact]
    public void RenderPage_UsesDeclaredAlternateAndReportsUnsupportedIccProfile() {
        byte[] unsupportedProfile = PdfIccProfiles.SrgbIec6196621;
        unsupportedProfile[16] = (byte)'C';
        unsupportedProfile[17] = (byte)'M';
        unsupportedProfile[18] = (byte)'Y';
        unsupportedProfile[19] = (byte)'K';
        byte[] pdf = BuildIccContentPdf(unsupportedProfile, "/N 3 /Alternate /DeviceRGB", "0.8 0.1 0.2 scn");

        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        OfficeImageExportResult result = page.ExportImage(OfficeImageExportFormat.Png);

        Assert.Equal(OfficeColor.FromRgb(204, 26, 51), Assert.Single(drawing.Shapes).Shape.FillColor);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.IccColorSpaceId);
    }

    [Fact]
    public void RenderPage_PassesUnsupportedIccComponentsDirectlyToDeclaredLabAlternate() {
        byte[] unsupportedProfile = PdfIccProfiles.SrgbIec6196621;
        unsupportedProfile[16] = (byte)'C';
        unsupportedProfile[17] = (byte)'M';
        unsupportedProfile[18] = (byte)'Y';
        unsupportedProfile[19] = (byte)'K';
        const string alternate = "[/Lab << /WhitePoint [1 1 1] /Range [-50 50 -25 25] >>]";
        byte[] pdf = BuildIccContentPdf(unsupportedProfile, "/N 3 /Alternate " + alternate, "50 40 -20 scn");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeColor expected = OfficeColorSpaceConverter.FromLab(50, 40, -20, 1, 1, 1);
        Assert.Equal(expected, Assert.Single(drawing.Shapes).Shape.FillColor);
    }

    [Fact]
    public void RenderPage_PreservesCalGrayParametersOnUnsupportedIccAlternate() {
        byte[] unsupportedProfile = PdfIccProfiles.SrgbIec6196621;
        unsupportedProfile[16] = (byte)'G';
        unsupportedProfile[17] = (byte)'R';
        unsupportedProfile[18] = (byte)'A';
        unsupportedProfile[19] = (byte)'Y';
        const string alternate = "[/CalGray << /WhitePoint [1 1 1] /Gamma 2 >>]";
        byte[] pdf = BuildIccContentPdf(unsupportedProfile, "/N 1 /Alternate " + alternate, "0.5 scn");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeColor expected = OfficeColorSpaceConverter.FromCalibratedGray(0.5, 1, 1, 1, 2);
        Assert.Equal(expected, Assert.Single(drawing.Shapes).Shape.FillColor);
    }

    [Fact]
    public void RenderPage_ClipsUnsupportedIccFallbackComponentsToRange() {
        byte[] unsupportedProfile = PdfIccProfiles.SrgbIec6196621;
        unsupportedProfile[16] = (byte)'G';
        unsupportedProfile[17] = (byte)'R';
        unsupportedProfile[18] = (byte)'A';
        unsupportedProfile[19] = (byte)'Y';
        const string alternate = "[/CalGray << /WhitePoint [0.9505 1 1.089] >>]";
        byte[] pdf = BuildIccContentPdf(
            unsupportedProfile,
            "/N 1 /Range [0.2 0.8] /Alternate " + alternate,
            "0 scn");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeColor expected = OfficeColorSpaceConverter.FromCalibratedGray(0.2, 0.9505, 1, 1.089, 1);
        Assert.Equal(expected, Assert.Single(drawing.Shapes).Shape.FillColor);
    }

    [Fact]
    public void RenderPage_ClipsImplicitUnsupportedIccFallbackComponentsToRange() {
        byte[] unsupportedProfile = PdfIccProfiles.SrgbIec6196621;
        unsupportedProfile[16] = (byte)'G';
        unsupportedProfile[17] = (byte)'R';
        unsupportedProfile[18] = (byte)'A';
        unsupportedProfile[19] = (byte)'Y';
        byte[] pdf = BuildIccContentPdf(
            unsupportedProfile,
            "/N 1 /Range [0.2 0.8]",
            "0 scn");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Equal(OfficeColor.FromRgb(51, 51, 51), Assert.Single(drawing.Shapes).Shape.FillColor);
    }

    [Fact]
    public void RenderPage_ScalesIndexedIccPaletteIntoDeclaredRange() {
        byte[] pdf = BuildIccContentPdf(
            PdfIccProfiles.SrgbIec6196621,
            "/N 3 /Range [-1 1 -1 1 -1 1]",
            "0 scn",
            colorSpaceResources: "/CsIcc [/Indexed [/ICCBased 5 0 R] 0 <000000>]");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.True(OfficeIccColorProfile.TryCreate(PdfIccProfiles.SrgbIec6196621, out OfficeIccColorProfile? profile));
        Assert.True(profile.TryConvert(new[] { 0D, 0D, 0D }, out OfficeColor expected));
        Assert.Equal(expected, Assert.Single(drawing.Shapes).Shape.FillColor);
    }

    [Fact]
    public void RenderPage_AppliesPaintTimeIntentToIndexedIccContentPalette() {
        byte[] profile = IccLutTestProfiles.CreateCmykLut8WithDistinctRelativeIntent();
        const string indexed = "/CsIcc [/Indexed [/ICCBased 5 0 R] 0 <00000000>]";
        OfficeColor perceptual = Assert.Single(PdfPageImageRenderer.RenderPage(BuildIccContentPdf(
            profile,
            "/N 4",
            string.Empty,
            colorSpaceResources: indexed,
            contentOverride: "/CsIcc cs 0 scn /Perceptual ri 0 0 10 10 re f")).Shapes).Shape.FillColor!.Value;
        OfficeColor relative = Assert.Single(PdfPageImageRenderer.RenderPage(BuildIccContentPdf(
            profile,
            "/N 4",
            string.Empty,
            colorSpaceResources: indexed,
            contentOverride: "/CsIcc cs 0 scn /RelativeColorimetric ri 0 0 10 10 re f")).Shapes).Shape.FillColor!.Value;

        Assert.NotEqual(perceptual, relative);
    }

    [Fact]
    public void RenderPage_UsesDeclaredSeparationAlternateForUnsupportedIccProfile() {
        byte[] pdf = BuildIccContentPdf(
            PdfIccProfiles.SrgbIec6196621,
            "/N 1 /Alternate [/Separation /Spot /DeviceRGB 7 0 R]",
            "1 scn",
            extraObjects: "7 0 obj\n<< /FunctionType 2 /Domain [0 1] /C0 [0 0 0] /C1 [0 1 0] /N 1 >>\nendobj\n");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Equal(OfficeColor.FromRgb(0, 255, 0), Assert.Single(drawing.Shapes).Shape.FillColor);
    }

    [Fact]
    public void RenderPage_FailsClosedWhenAuthoredIccAlternateCannotBeProjected() {
        byte[] unsupportedProfile = PdfIccProfiles.SrgbIec6196621;
        unsupportedProfile[16] = (byte)'G';
        unsupportedProfile[17] = (byte)'R';
        unsupportedProfile[18] = (byte)'A';
        unsupportedProfile[19] = (byte)'Y';
        byte[] pdf = BuildIccContentPdf(
            unsupportedProfile,
            "/N 1 /Alternate [/Separation /Spot /DeviceRGB 7 0 R]",
            "1 scn",
            extraObjects: "7 0 obj\n<< /FunctionType 3 /Domain [0 1] >>\nendobj\n");

        OfficeImageExportResult result = PdfReadDocument.Open(pdf).Pages[0].ExportImage(OfficeImageExportFormat.Png);

        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Fact]
    public void RenderPage_TreatsNullIccAlternateAsAbsent() {
        byte[] unsupportedProfile = PdfIccProfiles.SrgbIec6196621;
        unsupportedProfile[16] = (byte)'G';
        unsupportedProfile[17] = (byte)'R';
        unsupportedProfile[18] = (byte)'A';
        unsupportedProfile[19] = (byte)'Y';
        byte[] pdf = BuildIccContentPdf(
            unsupportedProfile,
            "/N 1 /Alternate null",
            "0.5 scn");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.InRange(Assert.Single(drawing.Shapes).Shape.FillColor!.Value.R, 126, 129);
    }

    [Theory]
    [InlineData("null", "")]
    [InlineData("7 0 R", "7 0 obj\n8 0 R\nendobj\n8 0 obj\nnull\nendobj\n")]
    public void RenderPage_TreatsNullIccRangeAsAbsent(string rangeValue, string extraObjects) {
        byte[] unsupportedProfile = PdfIccProfiles.SrgbIec6196621;
        unsupportedProfile[16] = (byte)'G';
        unsupportedProfile[17] = (byte)'R';
        unsupportedProfile[18] = (byte)'A';
        unsupportedProfile[19] = (byte)'Y';
        byte[] pdf = BuildIccContentPdf(
            unsupportedProfile,
            "/N 1 /Range " + rangeValue,
            "0.5 scn",
            extraObjects);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.InRange(Assert.Single(drawing.Shapes).Shape.FillColor!.Value.R, 126, 129);
    }

    [Theory]
    [InlineData("[/CalRGB << /WhitePoint [0.9505 1 1.089] /Gamma null /Matrix 7 0 R >>]", "0.5 0.5 0.5 scn")]
    [InlineData("[/CalGray << /WhitePoint [0.9505 1 1.089] /Gamma 7 0 R >>]", "0.5 scn")]
    [InlineData("[/Lab << /WhitePoint [0.9505 1 1.089] /Range 7 0 R >>]", "50 0 0 scn")]
    public void RenderPage_TreatsNullCalibratedOptionsAsAbsent(string colorSpace, string colorOperation) {
        byte[] pdf = BuildIccContentPdf(
            PdfIccProfiles.SrgbIec6196621,
            "/N 3",
            colorOperation,
            "7 0 obj\n8 0 R\nendobj\n8 0 obj\nnull\nendobj\n",
            colorSpaceName: "CsCal",
            colorSpaceResources: "/CsCal " + colorSpace);

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        PdfPageRenderResult result = Assert.Single(PdfPageImageRenderer.RenderPages(
            pdf,
            options: new PdfPageRenderOptions { Format = PdfPageRenderFormat.Svg, ContinueOnError = true }));

        Assert.Single(drawing.Shapes);
        Assert.DoesNotContain(
            result.CapabilityDiagnostics,
            diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId && diagnostic.Subject == "CsCal");
    }

    [Fact]
    public void RenderPage_ResolvesMultiHopIccProfileOperand() {
        byte[] pdf = BuildIccContentPdf(
            PdfIccProfiles.SrgbIec6196621,
            "/N 3",
            "1 0 0 scn",
            "7 0 obj\n8 0 R\nendobj\n8 0 obj\n5 0 R\nendobj\n",
            colorSpaceResources: "/CsIcc [/ICCBased 7 0 R]");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        Assert.Single(drawing.Shapes);
    }

    [Theory]
    [InlineData("CalGray", "<< /WhitePoint [0.9505 1 1.089] /Gamma 2 >>", "0.5 scn")]
    [InlineData("CalRGB", "<< /WhitePoint [0.9505 1 1.089] /Gamma [2 2 2] >>", "0.5 0.5 0.5 scn")]
    [InlineData("Lab", "<< /WhitePoint [0.9505 1 1.089] /Range [-100 100 -100 100] >>", "50 0 0 scn")]
    public void RenderPage_ResolvesMultiHopCalibratedDictionaries(
        string colorSpaceKind,
        string calibrationDictionary,
        string colorOperation) {
        byte[] pdf = BuildIccContentPdf(
            PdfIccProfiles.SrgbIec6196621,
            "/N 3",
            colorOperation,
            "7 0 obj\n8 0 R\nendobj\n8 0 obj\n" + calibrationDictionary + "\nendobj\n",
            colorSpaceName: "CsCal",
            colorSpaceResources: "/CsCal [/" + colorSpaceKind + " 7 0 R]");
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];

        Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Shapes);
        Assert.DoesNotContain(
            page.GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId && diagnostic.Subject == "CsCal");
    }

    [Fact]
    public void RenderPage_ResolvesMultiHopIccRangeEndpoints() {
        byte[] pdf = BuildIccContentPdf(
            PdfIccProfiles.SrgbIec6196621,
            "/N 3 /Range [7 0 R 1 0 1 0 1]",
            "1 0 0 scn",
            "7 0 obj\n8 0 R\nendobj\n8 0 obj\n0\nendobj\n");

        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];

        Assert.Single(PdfPageImageRenderer.RenderPage(pdf).Shapes);
        Assert.DoesNotContain(page.GetRenderCapabilityDiagnostics(), diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Fact]
    public void IccProfile_RejectsNonD50HeaderIlluminant() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        WriteS15Fixed16(profile, 68, 0.95047D);
        WriteS15Fixed16(profile, 72, 1D);
        WriteS15Fixed16(profile, 76, 1.08883D);

        Assert.False(OfficeIccColorProfile.TryCreate(profile, out _));
    }

    [Fact]
    public void IccProfile_RejectsDecreasingSampledToneCurve() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        (int offset, int length) = FindTag(profile, "rTRC");
        Assert.True(length >= 16);
        WriteUInt32(profile, offset, 0x63757276U);
        WriteUInt32(profile, offset + 8, 2U);
        profile[offset + 12] = 0xFF;
        profile[offset + 13] = 0xFF;
        profile[offset + 14] = 0x00;
        profile[offset + 15] = 0x00;

        Assert.False(OfficeIccColorProfile.TryCreate(profile, out _));
    }

    [Fact]
    public void IccProfile_RejectsDecreasingParametricToneCurve() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        (int offset, int length) = FindTag(profile, "rTRC");
        Assert.True(length >= 40);
        WriteUInt32(profile, offset, 0x70617261U);
        profile[offset + 8] = 0;
        profile[offset + 9] = 4;
        profile[offset + 10] = 0;
        profile[offset + 11] = 0;
        WriteS15Fixed16(profile, offset + 12, 2.4D);
        WriteS15Fixed16(profile, offset + 16, 0.94787D);
        WriteS15Fixed16(profile, offset + 20, 0.05213D);
        WriteS15Fixed16(profile, offset + 24, -1D);
        WriteS15Fixed16(profile, offset + 28, 0.04045D);
        WriteS15Fixed16(profile, offset + 32, 0D);
        WriteS15Fixed16(profile, offset + 36, 0D);

        Assert.False(OfficeIccColorProfile.TryCreate(profile, out _));
    }

    [Fact]
    public void IccProfile_RejectsParametricCurveThatDropsAtUnitBoundary() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        (int offset, int length) = FindTag(profile, "rTRC");
        Assert.True(length >= 32);
        WriteUInt32(profile, offset, 0x70617261U);
        profile[offset + 8] = 0;
        profile[offset + 9] = 3;
        profile[offset + 10] = 0;
        profile[offset + 11] = 0;
        WriteS15Fixed16(profile, offset + 12, 1D);
        WriteS15Fixed16(profile, offset + 16, 0.1D);
        WriteS15Fixed16(profile, offset + 20, 0D);
        WriteS15Fixed16(profile, offset + 24, 0.8D);
        WriteS15Fixed16(profile, offset + 28, 1D);

        Assert.False(OfficeIccColorProfile.TryCreate(profile, out _));
    }

    [Fact]
    public void IccProfile_AcceptsUnreachableLowerBranchAtZeroBoundary() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        (int offset, int length) = FindTag(profile, "rTRC");
        Assert.True(length >= 40);
        WriteUInt32(profile, offset, 0x70617261U);
        profile[offset + 8] = 0;
        profile[offset + 9] = 4;
        profile[offset + 10] = 0;
        profile[offset + 11] = 0;
        WriteS15Fixed16(profile, offset + 12, 1D);
        WriteS15Fixed16(profile, offset + 16, 1D);
        WriteS15Fixed16(profile, offset + 20, 0D);
        WriteS15Fixed16(profile, offset + 24, 1D);
        WriteS15Fixed16(profile, offset + 28, 0D);
        WriteS15Fixed16(profile, offset + 32, 0D);
        WriteS15Fixed16(profile, offset + 36, 0.5D);

        Assert.True(OfficeIccColorProfile.TryCreate(profile, out _));
    }

    [Fact]
    public void IccProfile_RejectsSingularRgbColorantMatrix() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        foreach (string tag in new[] { "rXYZ", "gXYZ", "bXYZ" }) {
            (int offset, _) = FindTag(profile, tag);
            WriteS15Fixed16(profile, offset + 8, 0D);
            WriteS15Fixed16(profile, offset + 12, 0D);
            WriteS15Fixed16(profile, offset + 16, 0D);
        }

        Assert.False(OfficeIccColorProfile.TryCreate(profile, out _));
    }

    [Fact]
    public void ExtractImages_ClipsImplicitCmykFallbackToIccRangeAfterExplicitDecode() {
        byte[] unsupportedProfile = PdfIccProfiles.SrgbIec6196621;
        unsupportedProfile[16] = (byte)'C';
        unsupportedProfile[17] = (byte)'M';
        unsupportedProfile[18] = (byte)'Y';
        unsupportedProfile[19] = (byte)'K';
        byte[] pdf = BuildIccImagePdf(
            unsupportedProfile,
            new byte[] { 0, 0, 0, 0 },
            "/N 4 /Range [0.2 0.8 0.2 0.8 0.2 0.8 0.2 0.8]",
            imageEntries: "/Decode [0 1 0 1 0 1 0 1]");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        Assert.True(OfficePngReader.TryDecode(Assert.Single(drawing.Images).Bytes, out OfficeRasterImage? raster));

        Assert.Equal(OfficeColor.FromRgb(153, 153, 153), raster!.GetPixel(0, 0));
    }

    [Theory]
    [InlineData("CalRGB", "Gamma")]
    [InlineData("CalRGB", "Matrix")]
    [InlineData("CalGray", "Gamma")]
    [InlineData("Lab", "Range")]
    public void ImageColorSpace_TreatsIndirectNullCalibratedOptionsAsAbsent(string name, string option) {
        PdfArray colorSpace = CreateCalibratedColorSpace(name, option, new PdfReference(7, 0));
        var objects = new Dictionary<int, PdfIndirectObject> {
            [7] = new PdfIndirectObject(7, 0, new PdfReference(8, 0)),
            [8] = new PdfIndirectObject(8, 0, PdfNull.Instance)
        };

        Assert.True(PdfImageColorSpaceNormalization.TryResolve(
            colorSpace,
            string.Empty,
            objects,
            PdfReadLimits.DefaultMaxDecodedStreamBytes,
            out _));
    }

    [Fact]
    public void ImageColorSpace_RejectsCyclicCalibratedOptionReference() {
        PdfArray colorSpace = CreateCalibratedColorSpace("CalGray", "Gamma", new PdfReference(7, 0));
        var objects = new Dictionary<int, PdfIndirectObject> {
            [7] = new PdfIndirectObject(7, 0, new PdfReference(8, 0)),
            [8] = new PdfIndirectObject(8, 0, new PdfReference(7, 0))
        };

        Assert.False(PdfImageColorSpaceNormalization.TryResolve(
            colorSpace,
            string.Empty,
            objects,
            PdfReadLimits.DefaultMaxDecodedStreamBytes,
            out _));
    }

#if NET8_0_OR_GREATER
    [Fact]
    public void SeparationImageConversionReusesTintOutputBuffers() {
        var function = new PdfDictionary();
        function.Items["FunctionType"] = new PdfNumber(2);
        function.Items["Domain"] = NumberArray(0, 1);
        function.Items["C0"] = NumberArray(0, 0, 0);
        function.Items["C1"] = NumberArray(0, 1, 0);
        function.Items["N"] = new PdfNumber(1);
        var colorSpace = new PdfArray();
        colorSpace.Items.Add(new PdfName("Separation"));
        colorSpace.Items.Add(new PdfName("Spot"));
        colorSpace.Items.Add(new PdfName("DeviceRGB"));
        colorSpace.Items.Add(function);
        Assert.True(PdfImageColorSpaceNormalization.TryResolve(
            colorSpace,
            string.Empty,
            new Dictionary<int, PdfIndirectObject>(),
            PdfReadLimits.DefaultMaxDecodedStreamBytes,
            out PdfImageColorSpaceNormalization normalization));
        PdfImageColorConversionBuffer conversionBuffer = normalization.CreateConversionBuffer();
        byte[] sample = { 255 };
        for (int index = 0; index < 32; index++) {
            Assert.True(normalization.TryConvertPixel(sample, 0, null, conversionBuffer, out _));
        }

        long before = GC.GetAllocatedBytesForCurrentThread();
        bool converted = true;
        OfficeColor color = OfficeColor.Black;
        for (int index = 0; index < 4096; index++) {
            converted &= normalization.TryConvertPixel(sample, 0, null, conversionBuffer, out color);
        }
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;

        Assert.True(converted);
        Assert.Equal(OfficeColor.FromRgb(0, 255, 0), color);
        Assert.InRange(allocated, 0, 1024);
    }
#endif

    [Fact]
    public void RenderPage_ClipsType2TintOutputsToDeclaredRange() {
        byte[] pdf = BuildIccContentPdf(
            PdfIccProfiles.SrgbIec6196621,
            "/N 1 /Alternate [/Separation /Spot [/Lab << /WhitePoint [1 1 1] >>] 7 0 R]",
            "1 scn",
            extraObjects: "7 0 obj\n<< /FunctionType 2 /Domain [0 1] /Range [0 1 0 1 0 1] /C0 [50 0 0] /C1 [50 0 0] /N 1 >>\nendobj\n");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeColor expected = OfficeColorSpaceConverter.FromLab(1D, 0D, 0D, 1D, 1D, 1D);
        Assert.Equal(expected, Assert.Single(drawing.Shapes).Shape.FillColor);
    }

    [Fact]
    public void ExtractImages_AppliesEmbeddedMatrixTrcProfileAndDefaultIccRangeDecode() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        SwapTagPayload(profile, "rXYZ", "bXYZ");
        byte[] pdf = BuildIccImagePdf(profile, new byte[] { 255, 0, 0 }, "/N 3 /Range [0 1 0 1 0 1]");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        OfficeColor pixel = raster!.GetPixel(0, 0);
        Assert.True(pixel.B > 240);
        Assert.True(pixel.R < 40);
    }

    [Fact]
    public void ExtractImages_ColorManagesDctSamplesThroughPdfIccColorSpace() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        SwapTagPayload(profile, "rXYZ", "bXYZ");
        byte[] jpeg = CreateSinglePixelJpeg(OfficeColor.Red);
        byte[] pdf = BuildIccImagePdf(
            profile,
            jpeg,
            "/N 3 /Range [0 1 0 1 0 1]",
            imageEntries: "/Filter /DCTDecode");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.Equal("DCTDecode", image.Filter);
        Assert.Equal("png", image.FileExtension);
        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        OfficeColor pixel = raster!.GetPixel(0, 0);
        Assert.True(pixel.B > 220);
        Assert.True(pixel.R < 40);
    }

    [Fact]
    public void ExtractImages_AppliesPdfRenderingIntentToDctIccSamples() {
        byte[] profile = IccLutTestProfiles.CreateRgbLut16WithDistinctRelativeIntent();
        Assert.True(OfficeIccColorProfile.TryCreate(profile, out OfficeIccColorProfile? parsedProfile));
        Assert.True(parsedProfile!.TryConvert(new[] { 0D, 0D, 0D }, OfficeIccRenderingIntent.Perceptual, out _));
        Assert.True(parsedProfile.TryConvert(new[] { 0D, 0D, 0D }, OfficeIccRenderingIntent.RelativeColorimetric, out _));
        byte[] jpeg = CreateSinglePixelJpeg(OfficeColor.FromRgb(1, 0, 0));
        Assert.True(OfficeJpegCodec.TryDecode(jpeg, out _));
        Assert.True(OfficeJpegCodec.TryDecodeColorComponents(
            jpeg,
            requestedColorTransform: null,
            usePdfColorTransformDefault: true,
            out byte[] decodedComponents,
            out int decodedWidth,
            out int decodedHeight,
            out int decodedComponentCount));
        Assert.Equal(3, decodedComponents.Length);
        Assert.Equal(1, decodedWidth);
        Assert.Equal(1, decodedHeight);
        Assert.Equal(3, decodedComponentCount);
        OfficeColor perceptual = ReadSinglePixel(BuildIccImagePdf(
            profile,
            jpeg,
            "/N 3",
            imageEntries: "/Filter /DCTDecode /Intent /Perceptual"));
        OfficeColor relative = ReadSinglePixel(BuildIccImagePdf(
            profile,
            jpeg,
            "/N 3",
            imageEntries: "/Filter /DCTDecode /Intent /RelativeColorimetric"));

        Assert.NotEqual(perceptual, relative);
    }

    [Fact]
    public void ExtractImages_ColorManagesCmykDctSamplesThroughPdfIccColorSpace() {
        byte[] jpeg = CreateSinglePixelCmykJpeg();
        Assert.True(OfficeJpegCodec.TryDecode(jpeg, out OfficeRasterImage? decodedJpeg));
        OfficeColor decodedPixel = decodedJpeg!.GetPixel(0, 0);
        Assert.InRange(decodedPixel.R, 220, 255);
        Assert.InRange(decodedPixel.G, 0, 35);
        Assert.InRange(decodedPixel.B, 0, 35);
        byte[] pdf = BuildIccImagePdf(
            IccLutTestProfiles.CreateCmykLut8(),
            jpeg,
            "/N 4",
            imageEntries: "/Filter /DCTDecode");

        OfficeColor pixel = ReadSinglePixel(pdf);

        Assert.InRange(pixel.R, 220, 255);
        Assert.InRange(pixel.G, 0, 35);
        Assert.InRange(pixel.B, 0, 35);
    }

    [Fact]
    public void ExtractImages_AppliesDecodeAndSoftMaskToColorManagedDctSamples() {
        byte[] jpeg = CreateSinglePixelJpeg(OfficeColor.Red);
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            jpeg,
            "/N 3",
            imageEntries: "/Filter /DCTDecode /Decode [1 0 1 0 1 0] /SMask 7 0 R",
            softMaskSample: 128,
            imageColorSpace: "/DeviceRGB");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(image.TransparencyMaskResolved);
        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        OfficeColor pixel = raster!.GetPixel(0, 0);
        Assert.InRange(pixel.R, 0, 35);
        Assert.InRange(pixel.G, 220, 255);
        Assert.InRange(pixel.B, 220, 255);
        Assert.Equal(128, pixel.A);
    }

    [Fact]
    public void ExtractImages_AppliesChainedIndirectSoftMaskToColorManagedDctSamples() {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            CreateSinglePixelJpeg(OfficeColor.Red),
            "/N 3",
            imageEntries: "/Filter /DCTDecode /SMask 8 0 R",
            softMaskSample: 128,
            imageColorSpace: "/DeviceRGB",
            extraObjects:
                "8 0 obj\n9 0 R\nendobj\n" +
                "9 0 obj\n7 0 R\nendobj\n");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(image.TransparencyMaskResolved);
        Assert.False(image.HasUnresolvedTransparencyMask);
        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        Assert.Equal(128, raster!.GetPixel(0, 0).A);
    }

    [Fact]
    public void ExtractImages_AppliesColorKeyMaskToDctSamplesBeforeColorConversion() {
        byte[] jpeg = CreateSinglePixelJpeg(OfficeColor.Red);
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            jpeg,
            "/N 3",
            imageEntries: "/Filter /DCTDecode /Mask [200 255 0 50 0 50]",
            imageColorSpace: "/DeviceRGB");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(image.TransparencyMaskResolved);
        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        Assert.Equal(0, raster!.GetPixel(0, 0).A);
    }

    [Fact]
    public void ExtractImages_AppliesAndReportsChainedIndirectColorKeyMaskForDct() {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            CreateSinglePixelJpeg(OfficeColor.Red),
            "/N 3",
            imageEntries: "/Filter /DCTDecode /Mask 7 0 R",
            imageColorSpace: "/DeviceRGB",
            extraObjects:
                "7 0 obj\n8 0 R\nendobj\n" +
                "8 0 obj\n[200 255 0 50 0 50]\nendobj\n");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(image.TransparencyMaskResolved);
        Assert.False(image.HasUnresolvedTransparencyMask);
        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        Assert.Equal(0, raster!.GetPixel(0, 0).A);
    }

    [Fact]
    public void ExtractImages_HonorsDctColorTransformAndIndirectNullableDecodeParms() {
        byte[] jpeg = CreateSinglePixelJpeg(OfficeColor.Red);
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        OfficeColor direct = ReadSinglePixel(BuildIccImagePdf(
            profile,
            jpeg,
            "/N 3",
            imageEntries: "/Filter /DCTDecode /DecodeParms << /ColorTransform 1 >>",
            imageColorSpace: "/DeviceRGB"));
        OfficeColor indirect = ReadSinglePixel(BuildIccImagePdf(
            profile,
            jpeg,
            "/N 3",
            imageEntries: "/Filter /DCTDecode /DecodeParms 7 0 R",
            imageColorSpace: "/DeviceRGB",
            extraObjects: "7 0 obj\n<< /ColorTransform 8 0 R >>\nendobj\n8 0 obj\n1\nendobj\n"));
        byte[] explicitNullPdf = BuildIccImagePdf(
            profile,
            jpeg,
            "/N 3",
            imageEntries: "/Filter /DCTDecode /DecodeParms << /ColorTransform null >>",
            imageColorSpace: "/DeviceRGB");
        PdfExtractedImage explicitNull = Assert.Single(PdfImageExtractor.ExtractImages(explicitNullPdf));
        OfficeColor untransformed = ReadSinglePixel(BuildIccImagePdf(
            profile,
            jpeg,
            "/N 3",
            imageEntries: "/Filter /DCTDecode /DecodeParms << /ColorTransform 0 >>",
            imageColorSpace: "/DeviceRGB"));

        Assert.InRange(direct.R, 220, 255);
        Assert.InRange(direct.G, 0, 35);
        Assert.InRange(direct.B, 0, 35);
        Assert.Equal(direct, indirect);
        Assert.Equal("jpg", explicitNull.FileExtension);
        Assert.Equal(jpeg, explicitNull.Bytes);
        Assert.NotEqual(direct, untransformed);
    }

    [Fact]
    public void ExtractImages_ResolvesChainedIndirectDctDeclarations() {
        byte[] jpeg = CreateSinglePixelJpeg(OfficeColor.Red);
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            jpeg,
            "/N 3",
            imageEntries: "/Filter 7 0 R /DecodeParms 9 0 R",
            imageColorSpace: "/DeviceRGB",
            extraObjects:
                "7 0 obj\n8 0 R\nendobj\n" +
                "8 0 obj\n/DCTDecode\nendobj\n" +
                "9 0 obj\n10 0 R\nendobj\n" +
                "10 0 obj\n<< /ColorTransform 11 0 R >>\nendobj\n" +
                "11 0 obj\n12 0 R\nendobj\n" +
                "12 0 obj\n1\nendobj\n");

        OfficeColor pixel = ReadSinglePixel(pdf);

        Assert.InRange(pixel.R, 220, 255);
        Assert.InRange(pixel.G, 0, 35);
        Assert.InRange(pixel.B, 0, 35);
    }

    [Theory]
    [InlineData("/Decode [0 1]")]
    [InlineData("/Mask [0 255]")]
    public void ExtractImages_FailsClosedForMalformedDctSampleDeclarations(string declaration) {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            CreateSinglePixelJpeg(OfficeColor.Red),
            "/N 3",
            imageEntries: "/Filter /DCTDecode " + declaration,
            imageColorSpace: "/DeviceRGB");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));
        OfficeImageExportResult result = PdfReadDocument.Open(pdf).Pages[0].ExportImage(OfficeImageExportFormat.Png);

        Assert.False(image.IsImageFile);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void JpegDecoder_DecodesBaselineAndProgressiveAdobeYcck(bool progressive) {
        byte[] jpeg = progressive ? CreateProgressiveYcckJpeg() : CreateBaselineYcckJpeg();

        _ = OfficeJpegCodec.Decode(jpeg);

        Assert.True(OfficeJpegCodec.TryDecodeColorComponents(
            jpeg,
            requestedColorTransform: null,
            usePdfColorTransformDefault: true,
            out byte[] components,
            out int width,
            out int height,
            out int componentCount));
        Assert.Equal(2, width);
        Assert.Equal(2, height);
        Assert.Equal(4, componentCount);
        Assert.Equal(16, components.Length);
        Assert.True(OfficeJpegCodec.TryDecode(jpeg, out OfficeRasterImage? raster));
        Assert.Equal(2, raster!.Width);
        Assert.Equal(2, raster.Height);
    }

    [Fact]
    public void ExtractImages_FailsClosedForCyclicDctDeclarations() {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            CreateSinglePixelJpeg(OfficeColor.Red),
            "/N 3",
            imageEntries: "/Filter 7 0 R",
            imageColorSpace: "/DeviceRGB",
            extraObjects:
                "7 0 obj\n8 0 R\nendobj\n" +
                "8 0 obj\n7 0 R\nendobj\n");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));
        OfficeImageExportResult result = PdfReadDocument.Open(pdf).Pages[0].ExportImage(OfficeImageExportFormat.Png);

        Assert.False(image.IsImageFile);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Fact]
    public void JpegDecoder_RejectsInvalidAndComponentIncompatibleAdobeTransforms() {
        byte[] rgbTransformTwo = AddAdobeTransform(CreateSinglePixelJpeg(OfficeColor.Red), 2);
        byte[] rgbTransformThree = AddAdobeTransform(CreateSinglePixelJpeg(OfficeColor.Red), 3);
        byte[] cmykTransformOne = WithAdobeTransform(CreateSinglePixelCmykJpeg(), 1);

        Assert.False(OfficeJpegCodec.TryDecode(rgbTransformTwo, out _));
        Assert.False(OfficeJpegCodec.TryDecode(rgbTransformThree, out _));
        Assert.False(OfficeJpegCodec.TryDecode(cmykTransformOne, out _));
        Assert.False(OfficeJpegCodec.TryDecodeColorComponents(
            rgbTransformThree,
            requestedColorTransform: 1,
            usePdfColorTransformDefault: true,
            out _,
            out _,
            out _,
            out _));
    }

    [Fact]
    public void ExtractImages_ColorManagesDctAfterSupportedPrefixFilter() {
        byte[] jpeg = CreateSinglePixelJpeg(OfficeColor.Red);
        byte[] encodedJpeg = Encoding.ASCII.GetBytes(Convert.ToHexString(jpeg) + ">");
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            encodedJpeg,
            "/N 3",
            imageEntries: "/Filter [/ASCIIHexDecode /DCTDecode] /DecodeParms [null << /ColorTransform 1 >>]",
            imageColorSpace: "/DeviceRGB");

        OfficeColor pixel = ReadSinglePixel(pdf);

        Assert.InRange(pixel.R, 220, 255);
        Assert.InRange(pixel.G, 0, 35);
        Assert.InRange(pixel.B, 0, 35);
    }

    [Fact]
    public void ExtractImages_PreservesDctPayloadAfterSupportedPrefixFilter() {
        byte[] jpeg = CreateSinglePixelJpeg(OfficeColor.Red);
        byte[] encodedJpeg = Encoding.ASCII.GetBytes(Convert.ToHexString(jpeg) + ">");
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            encodedJpeg,
            "/N 3",
            imageEntries: "/Filter [/ASCIIHexDecode /DCTDecode]",
            imageColorSpace: "/DeviceRGB");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.Equal("jpg", image.FileExtension);
        Assert.Equal(jpeg, image.Bytes);
    }

    [Fact]
    public void ExtractImages_AdobeMarkerOverridesPdfDctColorTransform() {
        byte[] jpeg = CreateSinglePixelCmykJpeg();
        byte[] profile = IccLutTestProfiles.CreateCmykLut8();
        OfficeColor defaultTransform = ReadSinglePixel(BuildIccImagePdf(
            profile,
            jpeg,
            "/N 4",
            imageEntries: "/Filter /DCTDecode"));
        OfficeColor requestedTransform = ReadSinglePixel(BuildIccImagePdf(
            profile,
            jpeg,
            "/N 4",
            imageEntries: "/Filter /DCTDecode /DecodeParms << /ColorTransform 1 >>"));

        Assert.Equal(defaultTransform, requestedTransform);
    }

    [Fact]
    public void ExtractImages_ProjectsGrayscaleDctSamplesThroughIndexedPalette() {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            CreateSinglePixelGrayscaleJpeg(),
            "/N 3",
            imageEntries: "/Filter /DCTDecode",
            imageColorSpace: "[/Indexed /DeviceRGB 1 <FF000000FF00>]");

        OfficeColor pixel = ReadSinglePixel(pdf);

        Assert.Equal(OfficeColor.FromRgb(0, 255, 0), pixel);
    }

    [Fact]
    public void ExtractImages_FailsClosedAndReportsMalformedColorManagedDct() {
        byte[] jpeg = CreateSinglePixelJpeg(OfficeColor.Red);
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            jpeg,
            "/N 3",
            imageEntries: "/Filter /DCTDecode /DecodeParms << /ColorTransform 2 >>",
            imageColorSpace: "/DeviceRGB");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));
        OfficeImageExportResult result = PdfReadDocument.Open(pdf).Pages[0].ExportImage(OfficeImageExportFormat.Png);

        Assert.False(image.IsImageFile);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Fact]
    public void ExtractImages_EnforcesDecodedStreamLimitForDctComponents() {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            CreateSolidJpeg(20, 1, OfficeColor.Red),
            "/N 3",
            imageEntries: "/Filter /DCTDecode /DecodeParms << /ColorTransform 1 >>",
            imageColorSpace: "/DeviceRGB",
            width: 20);
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions {
            Limits = new PdfReadLimits { MaxDecodedStreamBytes = 50 }
        });

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() => document.ExtractImages());

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(50, exception.Limit);
        Assert.Equal(60, exception.Actual);
    }

    [Fact]
    public void ExtractImages_FailsClosedAndReportsDctDimensionMismatch() {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            CreateSinglePixelJpeg(OfficeColor.Red),
            "/N 3",
            imageEntries: "/Filter /DCTDecode /DecodeParms << /ColorTransform 1 >>",
            imageColorSpace: "/DeviceRGB",
            width: 2);

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));
        OfficeImageExportResult result = PdfReadDocument.Open(pdf).Pages[0].ExportImage(OfficeImageExportFormat.Png);

        Assert.False(image.IsImageFile);
        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Fact]
    public void ExtractImages_AppliesEmbeddedCmykLut8Profile() {
        byte[] pdf = BuildIccImagePdf(
            IccLutTestProfiles.CreateCmykLut8(),
            new byte[] { 0, 255, 255, 0 },
            "/N 4");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        OfficeColor pixel = raster!.GetPixel(0, 0);
        Assert.InRange(pixel.R, 245, 255);
        Assert.InRange(pixel.G, 0, 15);
        Assert.InRange(pixel.B, 0, 15);
    }

    [Fact]
    public void ExtractImages_UsesImageRenderingIntentAndPdfRelativeDefault() {
        byte[] profile = IccLutTestProfiles.CreateCmykLut8WithDistinctRelativeIntent();
        byte[] samples = { 0, 0, 0, 0 };
        OfficeColor perceptual = ReadSinglePixel(BuildIccImagePdf(
            profile,
            samples,
            "/N 4",
            imageEntries: "/Intent /Perceptual"));
        OfficeColor relative = ReadSinglePixel(BuildIccImagePdf(
            profile,
            samples,
            "/N 4",
            imageEntries: "/Intent /RelativeColorimetric"));
        OfficeColor inheritedDefault = ReadSinglePixel(BuildIccImagePdf(profile, samples, "/N 4"));
        OfficeColor unknownDefault = ReadSinglePixel(BuildIccImagePdf(
            profile,
            samples,
            "/N 4",
            imageEntries: "/Intent /ProducerSpecific"));
        OfficeColor indirectPerceptual = ReadSinglePixel(BuildIccImagePdf(
            profile,
            samples,
            "/N 4",
            imageEntries: "/Intent 7 0 R",
            extraObjects: "7 0 obj\n/Perceptual\nendobj\n"));

        Assert.NotEqual(perceptual, relative);
        Assert.Equal(relative, inheritedDefault);
        Assert.Equal(relative, unknownDefault);
        Assert.Equal(perceptual, indirectPerceptual);
    }

    [Fact]
    public void ExtractImages_PropagatesRenderingIntentThroughIndexedIccPalette() {
        byte[] profile = IccLutTestProfiles.CreateCmykLut8WithDistinctRelativeIntent();
        const string indexed = "[/Indexed [/ICCBased 6 0 R] 0 <00000000>]";
        OfficeColor perceptual = ReadSinglePixel(BuildIccImagePdf(
            profile,
            new byte[] { 0 },
            "/N 4",
            imageEntries: "/Intent /Perceptual",
            imageColorSpace: indexed));
        OfficeColor relative = ReadSinglePixel(BuildIccImagePdf(
            profile,
            new byte[] { 0 },
            "/N 4",
            imageEntries: "/Intent /RelativeColorimetric",
            imageColorSpace: indexed));

        Assert.NotEqual(perceptual, relative);
    }

    [Fact]
    public void ExtractImages_PreservesDistinctInheritedIntentsForRepeatedImageResource() {
        byte[] profile = IccLutTestProfiles.CreateCmykLut8WithDistinctRelativeIntent();
        const string content =
            "q /Perceptual ri 40 0 0 40 40 80 cm /Im1 Do Q\n" +
            "q /RelativeColorimetric ri 40 0 0 40 100 80 cm /Im1 Do Q";
        byte[] pdf = BuildIccImagePdf(
            profile,
            new byte[] { 0, 0, 0, 0 },
            "/N 4",
            contentOperations: content);

        IReadOnlyList<PdfExtractedImage> images = PdfReadDocument.Open(pdf).Pages[0].GetImages();
        OfficeColor[] colors = images.Select(image => {
            Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
            return raster!.GetPixel(0, 0);
        }).ToArray();

        Assert.Equal(2, colors.Length);
        Assert.NotEqual(colors[0], colors[1]);
    }

    [Fact]
    public void ExtractImages_ExplicitImageIntentOverridesEachInheritedIntent() {
        byte[] profile = IccLutTestProfiles.CreateCmykLut8WithDistinctRelativeIntent();
        const string content =
            "q /Perceptual ri 40 0 0 40 40 80 cm /Im1 Do Q\n" +
            "q /RelativeColorimetric ri 40 0 0 40 100 80 cm /Im1 Do Q";
        byte[] pdf = BuildIccImagePdf(
            profile,
            new byte[] { 0, 0, 0, 0 },
            "/N 4",
            imageEntries: "/Intent /Perceptual",
            contentOperations: content);

        OfficeColor[] colors = PdfReadDocument.Open(pdf).Pages[0].GetImages().Select(image => {
            Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
            return raster!.GetPixel(0, 0);
        }).ToArray();

        Assert.Equal(2, colors.Length);
        Assert.Equal(colors[0], colors[1]);
    }

    [Fact]
    public void ExtractImages_AppliesEmbeddedCmykMabProfile() {
        byte[] pdf = BuildIccImagePdf(
            IccMabTestProfiles.CreateCmykLab8(),
            new byte[] { 0, 255, 255, 0 },
            "/N 4");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        OfficeColor pixel = raster!.GetPixel(0, 0);
        Assert.InRange(pixel.R, 245, 255);
        Assert.InRange(pixel.G, 0, 15);
        Assert.InRange(pixel.B, 0, 15);
    }

    [Fact]
    public void ExtractImages_AppliesCmykLutProfileThroughIndexedPalette() {
        byte[] pdf = BuildIccImagePdf(
            IccLutTestProfiles.CreateCmykLut8(),
            new byte[] { 0 },
            "/N 4",
            imageColorSpace: "[/Indexed [/ICCBased 6 0 R] 0 <00FFFF00>]");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        OfficeColor pixel = raster!.GetPixel(0, 0);
        Assert.InRange(pixel.R, 245, 255);
        Assert.InRange(pixel.G, 0, 15);
        Assert.InRange(pixel.B, 0, 15);
    }


    [Fact]
    public void ExtractImages_AppliesEmbeddedProfileBeforeSoftMaskAlpha() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        SwapTagPayload(profile, "rXYZ", "bXYZ");
        byte[] pdf = BuildIccImagePdf(
            profile,
            new byte[] { 255, 0, 0 },
            "/N 3",
            imageEntries: "/SMask 7 0 R",
            softMaskSample: 128);

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        OfficeColor pixel = raster!.GetPixel(0, 0);
        Assert.True(pixel.B > 240);
        Assert.True(pixel.R < 40);
        Assert.Equal(128, pixel.A);
    }

    [Fact]
    public void ExtractImages_ResolvesMultiHopSoftMaskStream() {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            new byte[] { 255, 0, 0 },
            "/N 3",
            imageEntries: "/SMask 8 0 R",
            softMaskSample: 128,
            extraObjects: "8 0 obj\n7 0 R\nendobj\n");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        Assert.Equal(128, raster!.GetPixel(0, 0).A);
        Assert.Equal("soft-mask", image.TransparencyMaskKind);
        Assert.True(image.TransparencyMaskResolved);
    }

    [Fact]
    public void ExtractImages_TreatsMultiHopNullSoftMaskAsNoMask() {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            new byte[] { 0 },
            "/N 3",
            imageEntries: "/SMask 8 0 R",
            imageColorSpace: "[/Indexed /DeviceRGB 0 <FF0000>]",
            extraObjects: "8 0 obj\n9 0 R\nendobj\n9 0 obj\nnull\nendobj\n");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        Assert.Equal(OfficeColor.Red, raster!.GetPixel(0, 0));
        Assert.Null(image.TransparencyMaskKind);
    }

    [Fact]
    public void ExtractImages_AppliesColorKeyToRawSamplesBeforeIccConversion() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        byte[] pdf = BuildIccImagePdf(
            profile,
            new byte[] { 255, 0, 0 },
            "/N 3",
            imageEntries: "/Mask [255 255 0 0 0 0]");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        Assert.Equal(0, raster!.GetPixel(0, 0).A);
    }

    [Fact]
    public void ExtractImages_AppliesMultiHopColorKeyMask() {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            new byte[] { 255, 0, 0 },
            "/N 3",
            imageEntries: "/Mask 7 0 R",
            extraObjects: "7 0 obj\n8 0 R\nendobj\n8 0 obj\n[255 255 0 0 0 0]\nendobj\n");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        Assert.Equal(0, raster!.GetPixel(0, 0).A);
        Assert.Equal("color-key-mask", image.TransparencyMaskKind);
        Assert.True(image.TransparencyMaskResolved);
    }

    [Fact]
    public void ColorKeyDeclarationFailsClosedForMissingAndCyclicReferences() {
        var dictionary = new PdfDictionary();
        dictionary.Items["Mask"] = new PdfReference(7, 0);

        Assert.False(PdfImageColorKeyMask.TryCreateDeclaration(
            dictionary,
            componentCount: 3,
            new Dictionary<int, PdfIndirectObject>(),
            out _));

        var objects = new Dictionary<int, PdfIndirectObject> {
            [7] = new PdfIndirectObject(7, 0, new PdfReference(8, 0)),
            [8] = new PdfIndirectObject(8, 0, new PdfReference(7, 0))
        };
        Assert.False(PdfImageColorKeyMask.TryCreateDeclaration(
            dictionary,
            componentCount: 3,
            objects,
            out _));
    }

    [Fact]
    public void ExtractImages_ReportsMalformedMultiHopColorKeyMaskAsUnresolved() {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            new byte[] { 255, 0, 0 },
            "/N 3",
            imageEntries: "/Mask 7 0 R",
            extraObjects: "7 0 obj\n8 0 R\nendobj\n8 0 obj\n[255 /Bad 0 0 0 0]\nendobj\n");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.False(image.IsImageFile);
        Assert.Equal("color-key-mask", image.TransparencyMaskKind);
        Assert.True(image.HasUnresolvedTransparencyMask);
    }

    [Theory]
    [InlineData(new double[] { 0, 0, 0, 0, 0, 0, 0 })]
    [InlineData(new double[] { 0.5, 1, 0, 0, 0, 0 })]
    [InlineData(new double[] { -1, 0, 0, 0, 0, 0 })]
    [InlineData(new double[] { 0, 256, 0, 0, 0, 0 })]
    [InlineData(new double[] { 2, 1, 0, 0, 0, 0 })]
    public void ColorKeyMask_RejectsMalformedNumericRanges(double[] values) {
        var dictionary = new PdfDictionary();
        dictionary.Items["Mask"] = NumberArray(values);

        Assert.Null(PdfImageColorKeyMask.Create(dictionary, 3, 8, new Dictionary<int, PdfIndirectObject>()));
    }

    [Fact]
    public void RenderPage_AdaptivelySamplesIccShadingBeforeSrgbInterpolation() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        OfficeLinearGradient gradient = Assert.Single(
            PdfPageImageRenderer.RenderPage(BuildIccShadingPdf(profile)).Shapes,
            item => item.Shape.FillGradient != null).Shape.FillGradient!;

        Assert.True(gradient.Stops.Count > 2);
        OfficeGradientStop midpoint = Assert.Single(
            gradient.Stops,
            stop => Math.Abs(stop.Offset - 0.5D) < 0.000001D);
        Assert.True(OfficeIccColorProfile.TryCreate(profile, out OfficeIccColorProfile? parsedProfile));
        Assert.True(parsedProfile!.TryConvert(new[] { 0.5D, 0.5D, 0.5D }, out OfficeColor expected));
        Assert.Equal(expected, midpoint.Color);
        Assert.NotEqual((byte)128, midpoint.Color.R);
    }

    [Fact]
    public void RenderPage_AdaptivelySamplesEveryIccStitchingSubfunction() {
        const string function = "<< /FunctionType 3 /Domain [0 1] /Functions [" +
            "<< /FunctionType 2 /Domain [0 1] /C0 [0 0 0] /C1 [1 1 1] /N 1 >> " +
            "<< /FunctionType 2 /Domain [0 1] /C0 [0 0 0] /C1 [1 1 1] /N 1 >>] " +
            "/Bounds [0.5] /Encode [0 1 0 1] >>";
        OfficeLinearGradient gradient = Assert.Single(
            PdfPageImageRenderer.RenderPage(BuildIccShadingPdf(PdfIccProfiles.SrgbIec6196621, function)).Shapes,
            item => item.Shape.FillGradient != null).Shape.FillGradient!;

        Assert.True(gradient.Stops.Count > 4);
        Assert.Contains(gradient.Stops, stop => stop.Offset > 0D && stop.Offset < 0.5D);
        Assert.Contains(gradient.Stops, stop => stop.Offset > 0.5D && stop.Offset < 1D);
    }

    [Fact]
    public void CalculatorTintFunctionRejectsUndecodableFilteredPayload() {
        var dictionary = new PdfDictionary();
        dictionary.Items["FunctionType"] = new PdfNumber(4);
        dictionary.Items["Domain"] = NumberArray(0, 1);
        dictionary.Items["Range"] = NumberArray(0, 1);
        dictionary.Items["Filter"] = new PdfName("FlateDecode");
        var stream = new PdfStream(dictionary, Encoding.ASCII.GetBytes("{}"));

        Assert.False(PdfColorSpaceFunctionResolver.TryCreateTintTransform(
            stream,
            inputCount: 1,
            outputCount: 1,
            new Dictionary<int, PdfIndirectObject>(),
            PdfReadLimits.DefaultMaxDecodedStreamBytes,
            out _));
    }

    [Fact]
    public void ExtractImages_AppliesIccBasedIndexedPalette() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        SwapTagPayload(profile, "rXYZ", "bXYZ");
        byte[] pdf = BuildIccImagePdf(
            profile,
            new byte[] { 0 },
            "/N 3",
            imageColorSpace: "[/Indexed [/ICCBased 6 0 R] 0 <FF0000>]");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        OfficeColor pixel = raster!.GetPixel(0, 0);
        Assert.True(pixel.B > 240);
        Assert.True(pixel.R < 40);
    }

    [Fact]
    public void ExtractImages_UsesCalRgbAlternateForUnsupportedIccProfile() {
        byte[] unsupportedProfile = PdfIccProfiles.SrgbIec6196621;
        unsupportedProfile[16] = (byte)'C';
        unsupportedProfile[17] = (byte)'M';
        unsupportedProfile[18] = (byte)'Y';
        unsupportedProfile[19] = (byte)'K';
        const string alternate = "[/CalRGB << /WhitePoint [0.9505 1 1.089] /Gamma [2 2 2] >>]";
        byte[] pdf = BuildIccImagePdf(
            unsupportedProfile,
            new byte[] { 128, 64, 32 },
            "/N 3 /Alternate " + alternate);

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        OfficeColor expected = OfficeColorSpaceConverter.FromCalibratedRgb(
            128D / 255D,
            64D / 255D,
            32D / 255D,
            0.9505D,
            1D,
            1.089D,
            new[] { 2D, 2D, 2D });
        Assert.Equal(expected, raster!.GetPixel(0, 0));
    }

    [Theory]
    [InlineData("null", "")]
    [InlineData("7 0 R", "7 0 obj\n8 0 R\nendobj\n8 0 obj\nnull\nendobj\n")]
    public void ExtractImages_TreatsNullIccAlternateAsAbsent(string alternateValue, string extraObjects) {
        byte[] unsupportedProfile = PdfIccProfiles.SrgbIec6196621;
        unsupportedProfile[16] = (byte)'G';
        unsupportedProfile[17] = (byte)'R';
        unsupportedProfile[18] = (byte)'A';
        unsupportedProfile[19] = (byte)'Y';
        byte[] pdf = BuildIccImagePdf(
            unsupportedProfile,
            new byte[] { 128 },
            "/N 1 /Alternate " + alternateValue,
            extraObjects: extraObjects);

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        Assert.InRange(raster!.GetPixel(0, 0).R, 126, 129);
    }

    [Theory]
    [InlineData("null", "")]
    [InlineData("7 0 R", "7 0 obj\n8 0 R\nendobj\n8 0 obj\nnull\nendobj\n")]
    public void ExtractImages_TreatsNullIccRangeAsAbsent(string rangeValue, string extraObjects) {
        byte[] unsupportedProfile = PdfIccProfiles.SrgbIec6196621;
        unsupportedProfile[16] = (byte)'G';
        unsupportedProfile[17] = (byte)'R';
        unsupportedProfile[18] = (byte)'A';
        unsupportedProfile[19] = (byte)'Y';
        byte[] pdf = BuildIccImagePdf(
            unsupportedProfile,
            new byte[] { 128 },
            "/N 1 /Range " + rangeValue,
            extraObjects: extraObjects);

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        Assert.InRange(raster!.GetPixel(0, 0).R, 126, 129);
    }

    [Theory]
    [InlineData("/Alternate 7 0 R")]
    [InlineData("/Range 7 0 R")]
    public void RenderPage_FailsClosedForCyclicOptionalIccImageDeclarations(string declaration) {
        byte[] unsupportedProfile = PdfIccProfiles.SrgbIec6196621;
        unsupportedProfile[16] = (byte)'G';
        unsupportedProfile[17] = (byte)'R';
        unsupportedProfile[18] = (byte)'A';
        unsupportedProfile[19] = (byte)'Y';
        byte[] pdf = BuildIccImagePdf(
            unsupportedProfile,
            new byte[] { 128 },
            "/N 1 " + declaration,
            extraObjects: "7 0 obj\n8 0 R\nendobj\n8 0 obj\n7 0 R\nendobj\n");
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];

        Assert.False(Assert.Single(PdfImageExtractor.ExtractImages(pdf)).IsImageFile);
        Assert.Contains(page.GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Fact]
    public void ExtractImages_UsesIccRangeAsDefaultDecodeBeforeLabFallback() {
        byte[] unsupportedProfile = PdfIccProfiles.SrgbIec6196621;
        unsupportedProfile[16] = (byte)'C';
        unsupportedProfile[17] = (byte)'M';
        unsupportedProfile[18] = (byte)'Y';
        unsupportedProfile[19] = (byte)'K';
        const string range = "/Range [0 100 -100 100 -100 100]";
        const string alternate = "[/Lab << /WhitePoint [0.9505 1 1.089] /Range [-100 100 -100 100] >>]";
        byte[] pdf = BuildIccImagePdf(
            unsupportedProfile,
            new byte[] { 128, 128, 64 },
            "/N 3 " + range + " /Alternate " + alternate);

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        OfficeColor expected = OfficeColorSpaceConverter.FromLab(
            128D / 255D * 100D,
            -100D + 128D / 255D * 200D,
            -100D + 64D / 255D * 200D,
            0.9505D,
            1D,
            1.089D);
        Assert.Equal(expected, raster!.GetPixel(0, 0));
    }

    [Fact]
    public void ExtractImages_PreservesExplicitIdentityDecodeForLabImage() {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            new byte[] { 128, 128, 64 },
            "/N 3",
            imageEntries: "/Decode [0 1 0 1 0 1]",
            imageColorSpace: "[/Lab << /WhitePoint [0.9505 1 1.089] /Range [-100 100 -100 100] >>]");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        OfficeColor expected = OfficeColorSpaceConverter.FromLab(
            128D / 255D,
            128D / 255D,
            64D / 255D,
            0.9505D,
            1D,
            1.089D);
        Assert.Equal(expected, raster!.GetPixel(0, 0));
    }

    [Fact]
    public void ExtractImages_ResolvesReferenceChainedDecodeArrayAndComponentsForLabImage() {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            new byte[] { 128, 128, 64 },
            "/N 3",
            imageEntries: "/Decode 7 0 R",
            imageColorSpace: "[/Lab << /WhitePoint [0.9505 1 1.089] /Range [-100 100 -100 100] >>]",
            extraObjects:
                "7 0 obj\n8 0 R\nendobj\n" +
                "8 0 obj\n[9 0 R 10 0 R 9 0 R 10 0 R 9 0 R 10 0 R]\nendobj\n" +
                "9 0 obj\n11 0 R\nendobj\n" +
                "10 0 obj\n12 0 R\nendobj\n" +
                "11 0 obj\n0\nendobj\n" +
                "12 0 obj\n1\nendobj\n");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        OfficeColor expected = OfficeColorSpaceConverter.FromLab(
            128D / 255D,
            128D / 255D,
            64D / 255D,
            0.9505D,
            1D,
            1.089D);
        Assert.Equal(expected, raster!.GetPixel(0, 0));
    }

    [Fact]
    public void ImageDecodeTransformRejectsReferenceCyclesInsteadOfApplyingDefaults() {
        var dictionary = new PdfDictionary();
        dictionary.Items["Decode"] = new PdfReference(7, 0);
        var objects = new Dictionary<int, PdfIndirectObject> {
            [7] = new PdfIndirectObject(7, 0, new PdfReference(8, 0)),
            [8] = new PdfIndirectObject(8, 0, new PdfReference(7, 0))
        };

        Assert.False(PdfImageDecodeTransform.TryCreateColor(dictionary, 1, objects, out _));
        Assert.False(PdfImageDecodeTransform.TryCreateIndexed(dictionary, objects, out _));
        Assert.False(PdfImageDecodeTransform.IsIdentityColorDecodeOrAbsent(dictionary, 1, objects));
    }

    [Fact]
    public void ExtractImages_PreservesExplicitIdentityDecodeForIccRange() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        byte[] pdf = BuildIccImagePdf(
            profile,
            new byte[] { 0, 0, 0 },
            "/N 3 /Range [-1 1 -1 1 -1 1]",
            imageEntries: "/Decode [0 1 0 1 0 1]");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        Assert.True(OfficeIccColorProfile.TryCreate(profile, out OfficeIccColorProfile? parsedProfile));
        Assert.True(parsedProfile.TryConvert(new[] { 0.5D, 0.5D, 0.5D }, out OfficeColor expected));
        Assert.Equal(expected, raster!.GetPixel(0, 0));
    }

    [Fact]
    public void ExtractImages_ClipsExplicitDecodeToUnsupportedIccFallbackRange() {
        byte[] unsupportedProfile = PdfIccProfiles.SrgbIec6196621;
        unsupportedProfile[16] = (byte)'G';
        unsupportedProfile[17] = (byte)'R';
        unsupportedProfile[18] = (byte)'A';
        unsupportedProfile[19] = (byte)'Y';
        byte[] pdf = BuildIccImagePdf(
            unsupportedProfile,
            new byte[] { 0 },
            "/N 1 /Range [0.2 0.8] /Alternate [/CalGray << /WhitePoint [0.9505 1 1.089] >>]",
            imageEntries: "/Decode [0 1]");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        OfficeColor expected = OfficeColorSpaceConverter.FromCalibratedGray(0.2, 0.9505, 1, 1.089, 1);
        Assert.Equal(expected, raster!.GetPixel(0, 0));
    }

    [Fact]
    public void ExtractImages_MapsDefaultDecodeThroughUnsupportedIccDeviceFallbackRange() {
        byte[] unsupportedProfile = PdfIccProfiles.SrgbIec6196621;
        unsupportedProfile[16] = (byte)'G';
        unsupportedProfile[17] = (byte)'R';
        unsupportedProfile[18] = (byte)'A';
        unsupportedProfile[19] = (byte)'Y';
        byte[] pdf = BuildIccImagePdf(
            unsupportedProfile,
            new byte[] { 128 },
            "/N 1 /Range [-1 1]");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        Assert.InRange(raster!.GetPixel(0, 0).R, 0, 2);
    }

    [Fact]
    public void ExtractImages_ClipsExplicitDecodeToUnsupportedIccDeviceFallbackRange() {
        byte[] unsupportedProfile = PdfIccProfiles.SrgbIec6196621;
        unsupportedProfile[16] = (byte)'G';
        unsupportedProfile[17] = (byte)'R';
        unsupportedProfile[18] = (byte)'A';
        unsupportedProfile[19] = (byte)'Y';
        byte[] pdf = BuildIccImagePdf(
            unsupportedProfile,
            new byte[] { 0 },
            "/N 1 /Range [0.2 0.8]",
            imageEntries: "/Decode [0 1]");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        Assert.InRange(raster!.GetPixel(0, 0).R, 50, 52);
    }

    [Fact]
    public void RenderDiagnostics_AcceptsManagedDctIccConversion() {
        var source = new OfficeRasterImage(1, 1, OfficeColor.Red);
        byte[] jpeg = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions {
            Quality = 100,
            Subsampling = OfficeJpegSubsampling.Y444
        });
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            jpeg,
            "/N 3",
            imageEntries: "/Filter /DCTDecode");
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];

        Assert.DoesNotContain(
            page.GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Fact]
    public void DctColorManagementAppliesNonIdentityDecodeForRgbIccFallback() {
        byte[] unsupportedProfile = PdfIccProfiles.SrgbIec6196621;
        unsupportedProfile[12] = (byte)'p';
        unsupportedProfile[13] = (byte)'r';
        unsupportedProfile[14] = (byte)'t';
        unsupportedProfile[15] = (byte)'r';
        var source = new OfficeRasterImage(1, 1, OfficeColor.Red);
        byte[] jpeg = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions {
            Quality = 100,
            Subsampling = OfficeJpegSubsampling.Y444
        });
        byte[] pdf = BuildIccImagePdf(
            unsupportedProfile,
            jpeg,
            "/N 3 /Alternate /DeviceRGB",
            imageEntries: "/Filter /DCTDecode /Decode [1 0 1 0 1 0]");
        PdfReadDocument document = PdfReadDocument.Open(pdf);

        Assert.DoesNotContain(
            document.Pages[0].GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));
        Assert.True(image.IsImageFile);
        Assert.Equal("image/png", image.MimeType);
    }

    [Fact]
    public void RenderDiagnostics_RejectsChainedDctBeforeClaimingIccProjection() {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            new byte[] { 0 },
            "/N 3",
            imageEntries: "/Filter [/ASCII85Decode /DCTDecode]");
        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];

        Assert.Contains(
            page.GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Fact]
    public void ImageCapabilityAcceptsIndirectSingletonDctForDeviceRgbPassThrough() {
        var image = new PdfDictionary();
        image.Items["Width"] = new PdfNumber(1);
        image.Items["Height"] = new PdfNumber(1);
        image.Items["BitsPerComponent"] = new PdfNumber(8);
        image.Items["ColorSpace"] = new PdfName("DeviceRGB");
        image.Items["Filter"] = new PdfReference(7, 0);
        var filterArray = new PdfArray();
        filterArray.Items.Add(new PdfName("DCTDecode"));
        var objects = new Dictionary<int, PdfIndirectObject> {
            [7] = new PdfIndirectObject(7, 0, new PdfReference(8, 0)),
            [8] = new PdfIndirectObject(8, 0, filterArray)
        };

        Assert.True(ResourceResolver.CanProjectImageColorSpace(
            image,
            resources: null,
            objects,
            PdfReadLimits.DefaultMaxDecodedStreamBytes));
    }

    [Fact]
    public void ImageCapabilityAcceptsIdentityAndIndirectNullDecodeForDctPassThrough() {
        var image = new PdfDictionary();
        image.Items["Width"] = new PdfNumber(1);
        image.Items["Height"] = new PdfNumber(1);
        image.Items["BitsPerComponent"] = new PdfNumber(8);
        image.Items["ColorSpace"] = new PdfName("DeviceRGB");
        image.Items["Filter"] = new PdfName("DCTDecode");
        var identityDecode = new PdfArray();
        for (int component = 0; component < 3; component++) {
            identityDecode.Items.Add(new PdfNumber(0));
            identityDecode.Items.Add(new PdfNumber(1));
        }
        image.Items["Decode"] = identityDecode;

        Assert.True(ResourceResolver.CanProjectImageColorSpace(
            image,
            resources: null,
            new Dictionary<int, PdfIndirectObject>(),
            PdfReadLimits.DefaultMaxDecodedStreamBytes));

        image.Items["Decode"] = new PdfReference(9, 0);
        var objects = new Dictionary<int, PdfIndirectObject> {
            [9] = new PdfIndirectObject(9, 0, new PdfReference(10, 0)),
            [10] = new PdfIndirectObject(10, 0, PdfNull.Instance)
        };
        Assert.True(ResourceResolver.CanProjectImageColorSpace(
            image,
            resources: null,
            objects,
            PdfReadLimits.DefaultMaxDecodedStreamBytes));
    }

    [Fact]
    public void ImageCapabilityRejectsMalformedDctDecode() {
        var image = new PdfDictionary();
        image.Items["Width"] = new PdfNumber(1);
        image.Items["Height"] = new PdfNumber(1);
        image.Items["BitsPerComponent"] = new PdfNumber(8);
        image.Items["ColorSpace"] = new PdfName("DeviceRGB");
        image.Items["Filter"] = new PdfName("DCTDecode");
        var malformedDecode = new PdfArray();
        malformedDecode.Items.Add(new PdfNumber(0));
        malformedDecode.Items.Add(new PdfNumber(1));
        image.Items["Decode"] = malformedDecode;

        Assert.False(ResourceResolver.CanProjectImageColorSpace(
            image,
            resources: null,
            new Dictionary<int, PdfIndirectObject>(),
            PdfReadLimits.DefaultMaxDecodedStreamBytes));
    }

    [Theory]
    [InlineData(false, 6)]
    [InlineData(true, 8)]
    public void CalGrayProjectionHonorsCallerLimitForExpandedScanlineBuffer(bool withSoftMask, int maxDecodedStreamBytes) {
        PdfStream stream = CreateCalGrayImageStream(width: 2, height: 1, withSoftMask);
        var objects = new Dictionary<int, PdfIndirectObject>();

        Assert.False(ResourceResolver.CanProjectImageColorSpace(
            stream.Dictionary,
            resources: null,
            objects,
            maxDecodedStreamBytes));

        PdfExtractedImage image = ResourceResolver.BuildExtractedImage(
            pageNumber: 1,
            resourceName: "Im1",
            objectNumber: 5,
            directStreamIdentity: 0,
            stream,
            objects,
            maxDecodedStreamBytes: maxDecodedStreamBytes);
        Assert.False(image.IsImageFile);
    }

    [Fact]
    public void ExtractImages_PreservesJpegThroughIndirectSingletonDctFilterChain() {
        var source = new OfficeRasterImage(1, 1, OfficeColor.Red);
        byte[] jpeg = OfficeJpegCodec.Encode(source, new OfficeJpegEncodeOptions {
            Quality = 100,
            Subsampling = OfficeJpegSubsampling.Y444
        });
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            jpeg,
            "/N 3",
            imageEntries: "/Filter 7 0 R",
            imageColorSpace: "/DeviceRGB",
            extraObjects: "7 0 obj\n8 0 R\nendobj\n8 0 obj\n[/DCTDecode]\nendobj\n");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(image.IsImageFile);
        Assert.Equal("jpg", image.FileExtension);
        Assert.Equal(jpeg, image.Bytes);
    }

    [Fact]
    public void IndexedFilteredLookupHonorsConfiguredDecodedStreamLimit() {
        var lookupDictionary = new PdfDictionary();
        lookupDictionary.Items["Filter"] = new PdfName("FlateDecode");
        var colorSpace = new PdfArray();
        colorSpace.Items.Add(new PdfName("Indexed"));
        colorSpace.Items.Add(new PdfName("DeviceRGB"));
        colorSpace.Items.Add(new PdfNumber(255));
        colorSpace.Items.Add(new PdfStream(lookupDictionary, Compress(new byte[768])));

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfIndexedImageNormalizer.CanNormalizeColorSpace(
                colorSpace,
                bitsPerComponent: 8,
                new Dictionary<int, PdfIndirectObject>(),
                maxDecodedStreamBytes: 100));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(100, exception.Limit);
    }

    [Fact]
    public void ExtractImages_ScalesIndexedLabPaletteIntoBaseRange() {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            new byte[] { 0 },
            "/N 3",
            imageColorSpace: "[/Indexed [/Lab << /WhitePoint [0.9505 1 1.089] /Range [-100 100 -100 100] >>] 0 <808040>]");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        OfficeColor expected = OfficeColorSpaceConverter.FromLab(
            128D / 255D * 100D,
            -100D + 128D / 255D * 200D,
            -100D + 64D / 255D * 200D,
            0.9505D,
            1D,
            1.089D);
        Assert.Equal(expected, raster!.GetPixel(0, 0));
    }

    [Fact]
    public void ExtractImages_AppliesIndexedCalGrayCalibration() {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            new byte[] { 0 },
            "/N 3",
            imageColorSpace: "[/Indexed [/CalGray << /WhitePoint [0.9505 1 1.089] /Gamma 2 >>] 0 <80>]");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        OfficeColor expected = OfficeColorSpaceConverter.FromCalibratedGray(
            128D / 255D,
            0.9505D,
            1D,
            1.089D,
            2D);
        Assert.Equal(expected, raster!.GetPixel(0, 0));
    }

    [Fact]
    public void ExtractImages_ConvertsDirectCalGrayToRgbScanlines() {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            new byte[] { 128 },
            "/N 3",
            imageColorSpace: "[/CalGray << /WhitePoint [0.9505 1 1.089] /Gamma 2 >>]");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        OfficeColor expected = OfficeColorSpaceConverter.FromCalibratedGray(
            128D / 255D,
            0.9505D,
            1D,
            1.089D,
            2D);
        Assert.Equal(expected, raster!.GetPixel(0, 0));
    }

    [Theory]
    [InlineData("/None")]
    [InlineData("null")]
    public void ExtractImages_TreatsIndexedEmptySoftMaskAsNoMask(string softMask) {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            new byte[] { 0 },
            "/N 3",
            imageEntries: "/SMask " + softMask,
            imageColorSpace: "[/Indexed /DeviceRGB 0 <FF0000>]");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        Assert.Equal(OfficeColor.Red, raster!.GetPixel(0, 0));
    }

    [Fact]
    public void ExtractImages_ClipsIndexedSampleToHighValueWhenDecodeIsOmitted() {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            new byte[] { 255 },
            "/N 3",
            imageColorSpace: "[/Indexed /DeviceRGB 1 <000000FF0000>]");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        Assert.Equal(OfficeColor.Red, raster!.GetPixel(0, 0));
    }

    [Fact]
    public void ExtractImages_AppliesExplicitIndexedDecodeBeforeHighValueClipping() {
        var palette = new byte[33];
        palette[5 * 3 + 1] = 255;
        palette[10 * 3] = 255;
        string lookup = BitConverter.ToString(palette).Replace("-", string.Empty);
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            new byte[] { 128 },
            "/N 3",
            imageEntries: "/Decode [0 10]",
            imageColorSpace: "[/Indexed /DeviceRGB 10 <" + lookup + ">]");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        Assert.Equal(OfficeColor.FromRgb(0, 255, 0), raster!.GetPixel(0, 0));
    }

    [Fact]
    public void ExtractImages_PreservesIccComponentCountForDeviceNAlternate() {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            new byte[] { 0, 255, 255, 0 },
            "/N 4 /Alternate [/DeviceN [/Cyan /Magenta /Yellow /Black] /DeviceCMYK 7 0 R]",
            extraObjects: "7 0 obj\n<< /FunctionType 4 /Domain [0 1 0 1 0 1 0 1] /Range [0 1 0 1 0 1 0 1] /Length 2 >>\nstream\n{}\nendstream\nendobj\n");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        Assert.Equal(OfficeColor.Red, raster!.GetPixel(0, 0));
    }

    [Theory]
    [InlineData(2, "/A /B")]
    [InlineData(5, "/A /B /C /D /E")]
    public void RenderPage_ReportsUnsupportedDeviceNImageWhenTintProgramCannotProject(
        int componentCount,
        string colorantNames) {
        byte[] samples = new byte[componentCount];
        string domain = string.Join(" ", Enumerable.Repeat("0 1", componentCount));
        string functionObject =
            "7 0 obj\n<< /FunctionType 4 /Domain [" + domain + "] /Range [0 1 0 1 0 1] /Length 2 >>\nstream\n{}\nendstream\nendobj\n";
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            samples,
            "/N " + componentCount.ToString(CultureInfo.InvariantCulture) +
            " /Alternate [/DeviceN [" + colorantNames + "] /DeviceRGB 7 0 R]",
            extraObjects: functionObject);

        PdfReadPage page = PdfReadDocument.Open(pdf).Pages[0];
        OfficeImageExportResult result = page.ExportImage(OfficeImageExportFormat.Png);

        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(4)]
    [InlineData(16)]
    public void RenderPage_ReportsCalibratedImageDepthsOutsideManagedProjection(int bitsPerComponent) {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            new byte[] { 0, 0 },
            "/N 3",
            imageColorSpace: "[/CalGray << /WhitePoint [0.9505 1 1.089] >>]",
            bitsPerComponent: bitsPerComponent);

        OfficeImageExportResult result = PdfReadDocument.Open(pdf).Pages[0].ExportImage(OfficeImageExportFormat.Png);

        Assert.Contains(result.Diagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Fact]
    public void ExtractImages_UsesDeclaredSeparationAlternateForUnsupportedIccProfile() {
        byte[] unsupportedProfile = PdfIccProfiles.SrgbIec6196621;
        unsupportedProfile[16] = (byte)'C';
        unsupportedProfile[17] = (byte)'M';
        unsupportedProfile[18] = (byte)'Y';
        unsupportedProfile[19] = (byte)'K';
        byte[] pdf = BuildIccImagePdf(
            unsupportedProfile,
            new byte[] { 255 },
            "/N 1 /Alternate [/Separation /Spot /DeviceRGB 7 0 R]",
            extraObjects: "7 0 obj\n<< /FunctionType 2 /Domain [0 1] /C0 [0 0 0] /C1 [0 1 0] /N 1 >>\nendobj\n");

        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));

        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        Assert.Equal(OfficeColor.FromRgb(0, 255, 0), raster!.GetPixel(0, 0));
    }

    [Fact]
    public void ExtractImages_UsesCallerDecodedStreamLimitForCompressedIccProfile() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        byte[] compressedProfile = Compress(profile);
        byte[] pdf = BuildIccImagePdf(
            compressedProfile,
            new byte[] { 255, 0, 0 },
            "/N 3 /Filter /FlateDecode");
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions {
            Limits = new PdfReadLimits { MaxDecodedStreamBytes = profile.Length - 1 }
        });

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfImageExtractor.ExtractImages(document));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(profile.Length - 1, exception.Limit);
    }

    [Fact]
    public void RenderPage_DoesNotDecodeUnusedIccColorSpaceResource() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        byte[] pdf = BuildIccContentPdf(
            Compress(profile),
            "/N 3 /Filter /FlateDecode",
            "1 0 0 scn",
            colorSpaceName: "CsRgb",
            colorSpaceResources: "/CsRgb /DeviceRGB /Unused [/ICCBased 5 0 R]");
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions {
            Limits = new PdfReadLimits { MaxDecodedStreamBytes = profile.Length - 1 }
        });

        OfficeImageExportResult result = document.Pages[0].ExportImage(OfficeImageExportFormat.Png);

        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.IccColorSpaceId);
    }

    [Fact]
    public void RenderDiagnostics_DoesNotScanNamedInlineImageSamplesAsOperators() {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        byte[] pdf = BuildNamedRawInlineImageWithUnusedIccPdf(Compress(profile));
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions {
            Limits = new PdfReadLimits { MaxDecodedStreamBytes = profile.Length - 1 }
        });

        PdfReadPage page = document.Pages[0];
        IReadOnlyList<PdfRenderCapabilityDiagnostic> diagnostics = page.GetRenderCapabilityDiagnostics();
        OfficeImageExportResult result = page.ExportImage(OfficeImageExportFormat.Png);

        Assert.DoesNotContain(diagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.IccColorSpaceId);
        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.IccColorSpaceId);
        Assert.NotEmpty(result.Bytes);
    }

    [Theory]
    [InlineData("shading")]
    [InlineData("shading-pattern")]
    [InlineData("tiling-pattern")]
    public void RenderPage_DoesNotDecodeIccProfileFromUnusedVisualResource(string resourceKind) {
        byte[] profile = PdfIccProfiles.SrgbIec6196621;
        byte[] pdf = BuildUnusedIccVisualResourcePdf(resourceKind, Compress(profile));
        PdfReadDocument document = PdfReadDocument.Open(pdf, new PdfReadOptions {
            Limits = new PdfReadLimits { MaxDecodedStreamBytes = profile.Length - 1 }
        });

        OfficeImageExportResult result = document.Pages[0].ExportImage(OfficeImageExportFormat.Png);

        Assert.DoesNotContain(result.Diagnostics, diagnostic => diagnostic.Code == PdfRenderCapabilities.IccColorSpaceId);
    }

    [Fact]
    public void Type2TintUsesDefaultsForDirectAndReferenceChainedNullEntries() {
        var function = new PdfDictionary();
        function.Items["FunctionType"] = new PdfNumber(2);
        function.Items["Domain"] = NumberArray(0, 1);
        function.Items["Range"] = new PdfReference(10, 0);
        function.Items["C0"] = PdfNull.Instance;
        function.Items["C1"] = new PdfReference(7, 0);
        function.Items["N"] = new PdfNumber(1);
        var objects = new Dictionary<int, PdfIndirectObject> {
            [7] = new PdfIndirectObject(7, 0, new PdfReference(8, 0)),
            [8] = new PdfIndirectObject(8, 0, PdfNull.Instance),
            [10] = new PdfIndirectObject(10, 0, new PdfReference(11, 0)),
            [11] = new PdfIndirectObject(11, 0, PdfNull.Instance)
        };

        Assert.True(PdfColorSpaceFunctionResolver.TryCreateTintTransform(
            function,
            inputCount: 1,
            outputCount: 1,
            objects,
            PdfReadLimits.DefaultMaxDecodedStreamBytes,
            out PdfColorSpaceTintTransform transform));
        var output = new double[1];

        Assert.True(transform(new[] { 1D }, output));
        Assert.Equal(1D, output[0]);
    }

    [Fact]
    public void RenderPage_RejectsNoneSeparationAcrossContentAndImageProjection() {
        const string function =
            "7 0 obj\n<< /FunctionType 2 /Domain [0 1] /C0 [0 0 0] /C1 [1 0 0] /N 1 >>\nendobj\n";
        byte[] contentPdf = BuildIccContentPdf(
            PdfIccProfiles.SrgbIec6196621,
            "/N 3",
            "1 scn",
            extraObjects: function,
            colorSpaceResources: "/CsIcc [/Separation /None /DeviceRGB 7 0 R]");
        byte[] imagePdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            new byte[] { 255 },
            "/N 3",
            imageColorSpace: "[/Separation /None /DeviceRGB 7 0 R]",
            extraObjects: function);

        Assert.Contains(
            PdfReadDocument.Open(contentPdf).Pages[0].GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
        Assert.Contains(
            PdfReadDocument.Open(imagePdf).Pages[0].GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
        Assert.False(Assert.Single(PdfImageExtractor.ExtractImages(imagePdf)).IsImageFile);
    }

    [Theory]
    [InlineData("[/CalGray << /WhitePoint [0.9505 1 1.089] /BlackPoint [0.1 0 0] >>]")]
    [InlineData("[/CalRGB << /WhitePoint [0.9505 1 1.089] /BlackPoint [0 0.1 0] >>]")]
    [InlineData("[/Lab << /WhitePoint [0.9505 1 1.089] /BlackPoint [0 0 0.1] >>]")]
    public void RenderPage_RejectsUnsupportedCalibratedBlackPointForImages(string colorSpace) {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            new byte[] { 128, 128, 128 },
            "/N 3",
            imageColorSpace: colorSpace);

        Assert.Contains(
            PdfReadDocument.Open(pdf).Pages[0].GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Theory]
    [InlineData("[/CalGray << /WhitePoint [0.9505 1 1.089] /BlackPoint [0.1 0 0] >>]", "0.5 scn")]
    [InlineData("[/CalRGB << /WhitePoint [0.9505 1 1.089] /BlackPoint [0 0.1 0] >>]", "0.5 0.5 0.5 scn")]
    [InlineData("[/Lab << /WhitePoint [0.9505 1 1.089] /BlackPoint [0 0 0.1] >>]", "50 0 0 scn")]
    public void RenderPage_RejectsUnsupportedCalibratedBlackPointForContent(string colorSpace, string operation) {
        byte[] pdf = BuildIccContentPdf(
            PdfIccProfiles.SrgbIec6196621,
            "/N 3",
            operation,
            colorSpaceResources: "/CsIcc " + colorSpace);

        Assert.Contains(
            PdfReadDocument.Open(pdf).Pages[0].GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Theory]
    [InlineData("[/CalGray << /WhitePoint [0.9505 2 1.089] >>]", "0.5 scn")]
    [InlineData("[/CalRGB << /WhitePoint [0.9505 2 1.089] >>]", "0.5 0.5 0.5 scn")]
    [InlineData("[/Lab << /WhitePoint [0.9505 2 1.089] >>]", "50 0 0 scn")]
    public void RenderPage_RejectsNonUnitCalibratedWhitePointForContent(string colorSpace, string operation) {
        byte[] pdf = BuildIccContentPdf(
            PdfIccProfiles.SrgbIec6196621,
            "/N 3",
            operation,
            colorSpaceResources: "/CsIcc " + colorSpace);

        Assert.Contains(
            PdfReadDocument.Open(pdf).Pages[0].GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Theory]
    [InlineData("[/CalGray << /WhitePoint [0.9505 2 1.089] >>]")]
    [InlineData("[/CalRGB << /WhitePoint [0.9505 2 1.089] >>]")]
    [InlineData("[/Lab << /WhitePoint [0.9505 2 1.089] >>]")]
    public void RenderPage_RejectsNonUnitCalibratedWhitePointForImages(string colorSpace) {
        byte[] pdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            new byte[] { 128, 128, 128 },
            "/N 3",
            imageColorSpace: colorSpace);

        Assert.Contains(
            PdfReadDocument.Open(pdf).Pages[0].GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Theory]
    [InlineData("[/Lab << /WhitePoint [0.9505 1 1.089] /Range [-129 127 -128 127] >>]")]
    [InlineData("[/Lab << /WhitePoint [0.9505 1 1.089] /Range [-128 128 -128 127] >>]")]
    [InlineData("[/Lab << /WhitePoint [0.9505 1 1.089] /Range [-128 127 -129 127] >>]")]
    [InlineData("[/Lab << /WhitePoint [0.9505 1 1.089] /Range [-128 127 -128 128] >>]")]
    public void RenderPage_RejectsLabRangesOutsideConverterDomainAcrossContentAndImages(string colorSpace) {
        byte[] contentPdf = BuildIccContentPdf(
            PdfIccProfiles.SrgbIec6196621,
            "/N 3",
            "50 0 0 scn",
            colorSpaceResources: "/CsIcc " + colorSpace);
        byte[] imagePdf = BuildIccImagePdf(
            PdfIccProfiles.SrgbIec6196621,
            new byte[] { 128, 128, 128 },
            "/N 3",
            imageColorSpace: colorSpace);

        Assert.Contains(PdfReadDocument.Open(contentPdf).Pages[0].GetRenderCapabilityDiagnostics(), diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
        Assert.Contains(PdfReadDocument.Open(imagePdf).Pages[0].GetRenderCapabilityDiagnostics(), diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Fact]
    public void CalibratedBlackPointAcceptsDefaultAndNullButRejectsReferenceCycles() {
        var calibration = new PdfDictionary();
        var objects = new Dictionary<int, PdfIndirectObject> {
            [7] = new PdfIndirectObject(7, 0, new PdfReference(8, 0)),
            [8] = new PdfIndirectObject(8, 0, PdfNull.Instance),
            [9] = new PdfIndirectObject(9, 0, new PdfReference(10, 0)),
            [10] = new PdfIndirectObject(10, 0, new PdfReference(9, 0))
        };

        calibration.Items["BlackPoint"] = NumberArray(0, 0, 0);
        Assert.True(PdfCalibratedColorSpaceSemantics.HasSupportedBlackPoint(calibration, objects));

        calibration.Items["BlackPoint"] = new PdfReference(7, 0);
        Assert.True(PdfCalibratedColorSpaceSemantics.HasSupportedBlackPoint(calibration, objects));

        calibration.Items["BlackPoint"] = new PdfReference(9, 0);
        Assert.False(PdfCalibratedColorSpaceSemantics.HasSupportedBlackPoint(calibration, objects));
    }

    [Fact]
    public void ManagedImageDecodeHonorsCallerLimitBeforeAllocatingSourcePixels() {
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Image");
        dictionary.Items["Width"] = new PdfNumber(2);
        dictionary.Items["Height"] = new PdfNumber(1);
        dictionary.Items["BitsPerComponent"] = new PdfNumber(8);
        dictionary.Items["ColorSpace"] = new PdfName("DeviceCMYK");
        dictionary.Items["Filter"] = new PdfName("FlateDecode");
        var stream = new PdfStream(dictionary, Compress(new byte[8]));
        var objects = new Dictionary<int, PdfIndirectObject>();

        Assert.False(ResourceResolver.CanProjectImageColorSpace(dictionary, null, objects, maxDecodedStreamBytes: 7));
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            ResourceResolver.BuildExtractedImage(
                pageNumber: 1,
                resourceName: "Im1",
                objectNumber: 5,
                directStreamIdentity: 0,
                stream,
                objects,
                maxDecodedStreamBytes: 7));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(7, exception.Limit);
    }

    [Fact]
    public void IndexedImageDecodeHonorsCallerLimitBeforeNormalizingPixels() {
        var dictionary = new PdfDictionary();
        dictionary.Items["Type"] = new PdfName("XObject");
        dictionary.Items["Subtype"] = new PdfName("Image");
        dictionary.Items["Width"] = new PdfNumber(8);
        dictionary.Items["Height"] = new PdfNumber(1);
        dictionary.Items["BitsPerComponent"] = new PdfNumber(8);
        dictionary.Items["Filter"] = new PdfName("FlateDecode");
        var colorSpace = new PdfArray();
        colorSpace.Items.Add(new PdfName("Indexed"));
        colorSpace.Items.Add(new PdfName("DeviceRGB"));
        colorSpace.Items.Add(new PdfNumber(1));
        colorSpace.Items.Add(new PdfStringObj(new byte[] { 255, 0, 0, 0, 255, 0 }));
        dictionary.Items["ColorSpace"] = colorSpace;
        var stream = new PdfStream(dictionary, Compress(new byte[8]));
        var objects = new Dictionary<int, PdfIndirectObject>();

        Assert.False(ResourceResolver.CanProjectImageColorSpace(dictionary, null, objects, maxDecodedStreamBytes: 7));
        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfIndexedImageNormalizer.TryBuildPngFile(
                colorSpace,
                width: 8,
                height: 1,
                bitsPerComponent: 8,
                stream,
                objects,
                maxDecodedStreamBytes: 7,
                out _));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(7, exception.Limit);
    }

    [Fact]
    public void IndexedSoftMaskDecodeHonorsCallerLimitBeforeNormalizingPixels() {
        var softMaskDictionary = new PdfDictionary();
        softMaskDictionary.Items["Type"] = new PdfName("XObject");
        softMaskDictionary.Items["Subtype"] = new PdfName("Image");
        softMaskDictionary.Items["Width"] = new PdfNumber(8);
        softMaskDictionary.Items["Height"] = new PdfNumber(1);
        softMaskDictionary.Items["BitsPerComponent"] = new PdfNumber(8);
        softMaskDictionary.Items["ColorSpace"] = new PdfName("DeviceGray");
        softMaskDictionary.Items["Filter"] = new PdfName("FlateDecode");
        var objects = new Dictionary<int, PdfIndirectObject> {
            [7] = new PdfIndirectObject(7, 0, new PdfStream(softMaskDictionary, Compress(new byte[8])))
        };

        var dictionary = new PdfDictionary();
        dictionary.Items["Width"] = new PdfNumber(8);
        dictionary.Items["Height"] = new PdfNumber(1);
        dictionary.Items["BitsPerComponent"] = new PdfNumber(1);
        dictionary.Items["Filter"] = new PdfName("FlateDecode");
        dictionary.Items["SMask"] = new PdfReference(7, 0);
        var colorSpace = new PdfArray();
        colorSpace.Items.Add(new PdfName("Indexed"));
        colorSpace.Items.Add(new PdfName("DeviceRGB"));
        colorSpace.Items.Add(new PdfNumber(1));
        colorSpace.Items.Add(new PdfStringObj(new byte[] { 255, 0, 0, 0, 255, 0 }));
        dictionary.Items["ColorSpace"] = colorSpace;
        var stream = new PdfStream(dictionary, Compress(new byte[] { 0 }));

        PdfReadLimitException exception = Assert.Throws<PdfReadLimitException>(() =>
            PdfIndexedImageNormalizer.TryBuildPngFile(
                colorSpace,
                width: 8,
                height: 1,
                bitsPerComponent: 1,
                stream,
                objects,
                maxDecodedStreamBytes: 7,
                out _));

        Assert.Equal(PdfReadLimitKind.DecodedStreamBytes, exception.Kind);
        Assert.Equal(7, exception.Limit);
    }

    private static byte[] BuildIccContentPdf(
        byte[] profile,
        string profileEntries,
        string colorOperation,
        string extraObjects = "",
        string colorSpaceName = "CsIcc",
        string colorSpaceResources = "/CsIcc [/ICCBased 5 0 R]",
        string extraResourceEntries = "",
        string? contentOverride = null) {
        string content = contentOverride ?? ("/" + colorSpaceName + " cs\n" + colorOperation + "\n40 80 70 40 re\nf");
        byte[] contentBytes = Encoding.ASCII.GetBytes(content);
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.4\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /ColorSpace << " + colorSpaceResources + " >> " + extraResourceEntries + " >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< " + profileEntries + " /Length " + profile.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(profile, 0, profile.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, extraObjects);
        WriteAscii(output, "trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildIccImagePdf(
        byte[] profile,
        byte[] imageSamples,
        string profileEntries,
        string imageEntries = "",
        byte? softMaskSample = null,
        string imageColorSpace = "[/ICCBased 6 0 R]",
        string extraObjects = "",
        int bitsPerComponent = 8,
        string? contentOperations = null,
        int width = 1,
        int height = 1) {
        byte[] contentBytes = Encoding.ASCII.GetBytes(contentOperations ?? "q\n40 0 0 40 40 80 cm\n/Im1 Do\nQ");
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.4\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /Type /XObject /Subtype /Image /Width " + width.ToString(CultureInfo.InvariantCulture) + " /Height " + height.ToString(CultureInfo.InvariantCulture) + " /BitsPerComponent " + bitsPerComponent.ToString(CultureInfo.InvariantCulture) + " /ColorSpace " + imageColorSpace + " " + imageEntries + " /Length " + imageSamples.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(imageSamples, 0, imageSamples.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "6 0 obj\n<< " + profileEntries + " /Length " + profile.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(profile, 0, profile.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        if (softMaskSample.HasValue) {
            WriteAscii(output, "7 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /BitsPerComponent 8 /ColorSpace /DeviceGray /Length 1 >>\nstream\n");
            output.WriteByte(softMaskSample.Value);
            WriteAscii(output, "\nendstream\nendobj\n");
        }
        WriteAscii(output, extraObjects);
        WriteAscii(output, "trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildIccImageMaskPdf(
        byte[] profile,
        string profileEntries,
        string contentOperations) {
        byte[] contentBytes = Encoding.ASCII.GetBytes(contentOperations);
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.4\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /ColorSpace << /CsIcc [/ICCBased 6 0 R] >> /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /ImageMask true /BitsPerComponent 1 /Length 1 >>\nstream\n");
        output.WriteByte(0x80);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "6 0 obj\n<< " + profileEntries + " /Length " + profile.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(profile, 0, profile.Length);
        WriteAscii(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildIccShadingPdf(byte[] profile, string? function = null) {
        function ??= "<< /FunctionType 2 /Domain [0 1] /C0 [0 0 0] /C1 [1 1 1] /N 1 >>";
        byte[] contentBytes = Encoding.ASCII.GetBytes("20 80 120 40 re\nW\nn\n/Sh1 sh");
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.4\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /Shading << /Sh1 5 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /ShadingType 2 /ColorSpace [/ICCBased 6 0 R] /Coords [20 80 140 80] /Function " + function + " /Extend [true true] >>\nendobj\n");
        WriteAscii(output, "6 0 obj\n<< /N 3 /Length " + profile.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(profile, 0, profile.Length);
        WriteAscii(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildUnusedIccVisualResourcePdf(string resourceKind, byte[] compressedProfile) {
        byte[] contentBytes = Encoding.ASCII.GetBytes("1 0 0 rg\n40 80 70 40 re\nf");
        string pageResources;
        string visualObjects;
        int profileObjectNumber;
        switch (resourceKind) {
            case "shading":
                pageResources = "/Shading << /Unused 6 0 R >>";
                visualObjects =
                    "6 0 obj\n<< /ShadingType 2 /ColorSpace [/ICCBased 8 0 R] /Coords [0 0 1 1] /Function 7 0 R >>\nendobj\n" +
                    "7 0 obj\n<< /FunctionType 2 /Domain [0 1] /C0 [0 0 0] /C1 [1 1 1] /N 1 >>\nendobj\n";
                profileObjectNumber = 8;
                break;
            case "shading-pattern":
                pageResources = "/Pattern << /Unused 6 0 R >>";
                visualObjects =
                    "6 0 obj\n<< /Type /Pattern /PatternType 2 /Shading 7 0 R >>\nendobj\n" +
                    "7 0 obj\n<< /ShadingType 2 /ColorSpace [/ICCBased 9 0 R] /Coords [0 0 1 1] /Function 8 0 R >>\nendobj\n" +
                    "8 0 obj\n<< /FunctionType 2 /Domain [0 1] /C0 [0 0 0] /C1 [1 1 1] /N 1 >>\nendobj\n";
                profileObjectNumber = 9;
                break;
            case "tiling-pattern":
                const string tileContent = "/TileIcc cs\n1 0 0 scn\n0 0 1 1 re\nf";
                pageResources = "/Pattern << /Unused 6 0 R >>";
                visualObjects =
                    "6 0 obj\n<< /Type /Pattern /PatternType 1 /PaintType 1 /TilingType 1 /BBox [0 0 1 1] /XStep 1 /YStep 1 " +
                    "/Resources << /ColorSpace << /TileIcc [/ICCBased 8 0 R] >> >> /Length " +
                    Encoding.ASCII.GetByteCount(tileContent).ToString(CultureInfo.InvariantCulture) +
                    " >>\nstream\n" + tileContent + "\nendstream\nendobj\n";
                profileObjectNumber = 8;
                break;
            default:
                throw new ArgumentOutOfRangeException(nameof(resourceKind));
        }

        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.4\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << " + pageResources + " >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, visualObjects);
        WriteAscii(output, profileObjectNumber.ToString(CultureInfo.InvariantCulture) + " 0 obj\n<< /N 3 /Filter /FlateDecode /Length " + compressedProfile.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(compressedProfile, 0, compressedProfile.Length);
        WriteAscii(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildNamedRawInlineImageWithUnusedIccPdf(byte[] compressedProfile) {
        using var content = new MemoryStream();
        WriteAscii(content, "q\n20 0 0 20 40 80 cm\nBI\n/W 5\n/H 1\n/CS /CsRgb\n/BPC 8\nID\n");
        byte[] samples = Encoding.ASCII.GetBytes("ABCDE/Unused cs");
        content.Write(samples, 0, samples.Length);
        WriteAscii(content, "\nEI\nQ");
        byte[] contentBytes = content.ToArray();

        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.4\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /ColorSpace << /CsRgb /DeviceRGB /Unused [/ICCBased 5 0 R] >> >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /N 3 /Filter /FlateDecode /Length " + compressedProfile.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(compressedProfile, 0, compressedProfile.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "trailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static PdfStream CreateCalGrayImageStream(int width, int height, bool withSoftMask) {
        var whitePoint = new PdfArray();
        whitePoint.Items.Add(new PdfNumber(0.9505));
        whitePoint.Items.Add(new PdfNumber(1));
        whitePoint.Items.Add(new PdfNumber(1.089));
        var calGrayDictionary = new PdfDictionary();
        calGrayDictionary.Items["WhitePoint"] = whitePoint;
        var colorSpace = new PdfArray();
        colorSpace.Items.Add(new PdfName("CalGray"));
        colorSpace.Items.Add(calGrayDictionary);

        var imageDictionary = new PdfDictionary();
        imageDictionary.Items["Type"] = new PdfName("XObject");
        imageDictionary.Items["Subtype"] = new PdfName("Image");
        imageDictionary.Items["Width"] = new PdfNumber(width);
        imageDictionary.Items["Height"] = new PdfNumber(height);
        imageDictionary.Items["BitsPerComponent"] = new PdfNumber(8);
        imageDictionary.Items["ColorSpace"] = colorSpace;
        if (withSoftMask) {
            var softMaskDictionary = new PdfDictionary();
            softMaskDictionary.Items["Type"] = new PdfName("XObject");
            softMaskDictionary.Items["Subtype"] = new PdfName("Image");
            softMaskDictionary.Items["Width"] = new PdfNumber(width);
            softMaskDictionary.Items["Height"] = new PdfNumber(height);
            softMaskDictionary.Items["BitsPerComponent"] = new PdfNumber(8);
            softMaskDictionary.Items["ColorSpace"] = new PdfName("DeviceGray");
            imageDictionary.Items["SMask"] = new PdfStream(softMaskDictionary, new byte[width * height]);
        }

        return new PdfStream(imageDictionary, new byte[width * height]);
    }

    private static PdfArray NumberArray(params double[] values) {
        var array = new PdfArray();
        for (int index = 0; index < values.Length; index++) array.Items.Add(new PdfNumber(values[index]));
        return array;
    }

    private static PdfArray CreateCalibratedColorSpace(string name, string option, PdfObject optionValue) {
        var calibration = new PdfDictionary();
        calibration.Items["WhitePoint"] = NumberArray(0.9505, 1, 1.089);
        calibration.Items[option] = optionValue;
        var colorSpace = new PdfArray();
        colorSpace.Items.Add(new PdfName(name));
        colorSpace.Items.Add(calibration);
        return colorSpace;
    }

    private static byte[] Compress(byte[] bytes) => OfficeZlibCodec.Compress(bytes);

    private static OfficeColor ReadSinglePixel(byte[] pdf) {
        PdfExtractedImage image = Assert.Single(PdfImageExtractor.ExtractImages(pdf));
        Assert.True(image.IsImageFile, "The extracted image payload was not normalized into an image file.");
        Assert.Equal("png", image.FileExtension);
        Assert.True(OfficePngReader.TryDecode(image.Bytes, out OfficeRasterImage? raster));
        return raster!.GetPixel(0, 0);
    }

    private static byte[] CreateSinglePixelJpeg(OfficeColor color) =>
        CreateSolidJpeg(1, 1, color);

    private static byte[] CreateSolidJpeg(int width, int height, OfficeColor color) {
        byte[] rgba = new byte[checked(width * height * 4)];
        for (int offset = 0; offset < rgba.Length; offset += 4) {
            rgba[offset] = color.R;
            rgba[offset + 1] = color.G;
            rgba[offset + 2] = color.B;
            rgba[offset + 3] = 255;
        }
        return
        OfficeJpegCodec.Encode(
            OfficeRasterImage.FromRgba32(width, height, rgba),
            new OfficeJpegEncodeOptions {
                Quality = 100,
                Subsampling = OfficeJpegSubsampling.Y444
            });
    }

    private static byte[] CreateSinglePixelCmykJpeg() => Convert.FromBase64String(
        "/9j/7gAOQWRvYmUAZAAAAAAA/9sAQwABAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEB/8AAFAgAAQABBEMRAE0RAFkRAEsRAP/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/aAA4EQwBNAFkASwAAPwD+/iv8/wDr/P8A6/v4r//Z");

    private static byte[] CreateBaselineYcckJpeg() => Convert.FromBase64String(
        "/9j/7gAOQWRvYmUAZAAAAAAC/9sAQwABAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEB/9sAQwEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEB/8AAFAgAAgACBAEiAAIRAQMRAQQiAP/EAB8AAAEFAQEBAQEBAAAAAAAAAAABAgMEBQYHCAkKC//EALUQAAIBAwMCBAMFBQQEAAABfQECAwAEEQUSITFBBhNRYQcicRQygZGhCCNCscEVUtHwJDNicoIJChYXGBkaJSYnKCkqNDU2Nzg5OkNERUZHSElKU1RVVldYWVpjZGVmZ2hpanN0dXZ3eHl6g4SFhoeIiYqSk5SVlpeYmZqio6Slpqeoqaqys7S1tre4ubrCw8TFxsfIycrS09TV1tfY2drh4uPk5ebn6Onq8fLz9PX29/j5+v/EAB8BAAMBAQEBAQEBAQEAAAAAAAABAgMEBQYHCAkKC//EALURAAIBAgQEAwQHBQQEAAECdwABAgMRBAUhMQYSQVEHYXETIjKBCBRCkaGxwQkjM1LwFWJy0QoWJDThJfEXGBkaJicoKSo1Njc4OTpDREVGR0hJSlNUVVZXWFlaY2RlZmdoaWpzdHV2d3h5eoKDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uLj5OXm5+jp6vLz9PX29/j5+v/aAA4EAQACEQMRBAAAPwD+1LSf2Qf2TNZ0rTNX1f8AZe/Z21XVtV0+y1LVNU1L4J/DW+1HUtRvraO6vb+/vbrwzLc3l7eXMstxdXVxLJPcTySTTSPI7MSiiv8ASbgX/kiODf8AsleHv/VRgz/Xfw1/5N1wB/2RXCn/AKostP6Gv+CZP/BMn/gm349/4Jt/8E+fHXjr/gnz+xB408beNP2IP2UPFnjHxj4s/ZQ+A3iPxV4s8VeI/gN4B1jxD4m8TeIdY8A3mr694g13V7y81TWdZ1S8utR1TUbq5vr65nuZ5ZWKKK//2Q==");

    private static byte[] CreateProgressiveYcckJpeg() => Convert.FromBase64String(
        "/9j/7gAOQWRvYmUAZAAAAAAC/9sAQwABAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEB/9sAQwEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEB/8IAFAgAAgACBAEiAAIRAQMRAQQiAP/EABUAAQEAAAAAAAAAAAAAAAAAAAAJ/8QAFAEBAAAAAAAAAAAAAAAAAAAACP/aAA4EAQACEAMQBAAAAAG1ISa7oaP/xAAVEAEBAAAAAAAAAAAAAAAAAAADBv/aAAgBAQABBQIpCTYv/8QAFREBAQAAAAAAAAAAAAAAAAAABDT/2gAIAQIBAQUCDF//xAAVEQEBAAAAAAAAAAAAAAAAAAAFNf/aAAgBAwEBBQI2d//EABUQAQEAAAAAAAAAAAAAAAAAAAcE/9oACAEEAAEFAjIyNrzb/8QAGxAAAwACAwAAAAAAAAAAAAAAAQIDBQYABBH/2gAIAQEABj8CnWur67WtZpStaYTGvSlHUM7u7dYs7uxLMzEliSSfef/EABgRAAIDAAAAAAAAAAAAAAAAAAAEAnSz/9oACAECAQY/Ak6q+UD/xAAYEQACAwAAAAAAAAAAAAAAAAAABAN0s//aAAgBAwEGPwJCkrhGf//EABsQAAMAAwEBAAAAAAAAAAAAAAIDBAEFBgcA/9oACAEEAAY/AvPrrvPuIttt4jlK7LK+U0NFVdVGhgdRTTQ6A2voe0za5zTJjWERmWSznP3/xAAVEAEBAAAAAAAAAAAAAAAAAAABAP/aAAgBAQABPyF5xEtS2AMUyW//xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oACAECAQE/IXr/xAAUEQEAAAAAAAAAAAAAAAAAAAAA/9oACAEDAQE/IVX/xAAUEAEAAAAAAAAAAAAAAAAAAAAB/9oACAEEAAE/IXH1hbTLICLU/9oADgQBAAIAAwAEAAAAEAQ//8QAFBABAAAAAAAAAAAAAAAAAAAAAf/aAAgBAQABPxBP+oCp5rhUFP/EABQRAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQIBAT8QJv/EABQRAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQMBAT8QPv/EABQQAQAAAAAAAAAAAAAAAAAAAAD/2gAIAQQAAT8QCCNuMkTGKn7/2Q==");

    private static byte[] CreateSinglePixelGrayscaleJpeg() => Convert.FromBase64String(
        "/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQEBAQH/wAALCAABAAEBAREA/8QAHwAAAQUBAQEBAQEAAAAAAAAAAAECAwQFBgcICQoL/8QAtRAAAgEDAwIEAwUFBAQAAAF9AQIDAAQRBRIhMUEGE1FhByJxFDKBkaEII0KxwRVS0fAkM2JyggkKFhcYGRolJicoKSo0NTY3ODk6Q0RFRkdISUpTVFVWV1hZWmNkZWZnaGlqc3R1dnd4eXqDhIWGh4iJipKTlJWWl5iZmqKjpKWmp6ipqrKztLW2t7i5usLDxMXGx8jJytLT1NXW19jZ2uHi4+Tl5ufo6erx8vP09fb3+Pn6/9oACAEBAAA/AP7+K//Z");

    private static byte[] AddAdobeTransform(byte[] jpeg, byte transform) {
        byte[] marker = {
            0xFF, 0xEE, 0x00, 0x0E,
            (byte)'A', (byte)'d', (byte)'o', (byte)'b', (byte)'e',
            0x00, 0x64,
            0x00, 0x00,
            0x00, 0x00,
            transform
        };
        var result = new byte[jpeg.Length + marker.Length];
        Buffer.BlockCopy(jpeg, 0, result, 0, 2);
        Buffer.BlockCopy(marker, 0, result, 2, marker.Length);
        Buffer.BlockCopy(jpeg, 2, result, 2 + marker.Length, jpeg.Length - 2);
        return result;
    }

    private static byte[] WithAdobeTransform(byte[] jpeg, byte transform) {
        byte[] result = (byte[])jpeg.Clone();
        for (int index = 2; index <= result.Length - 16; index++) {
            if (result[index] != 0xFF || result[index + 1] != 0xEE ||
                result[index + 4] != (byte)'A' || result[index + 5] != (byte)'d' ||
                result[index + 6] != (byte)'o' || result[index + 7] != (byte)'b' ||
                result[index + 8] != (byte)'e') continue;
            result[index + 15] = transform;
            return result;
        }
        throw new InvalidOperationException("Adobe APP14 marker was not found.");
    }

    private static OfficeColor ReadSingleRenderedImagePixel(byte[] pdf) {
        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);
        Assert.True(OfficePngReader.TryDecode(Assert.Single(drawing.Images).Bytes, out OfficeRasterImage? raster));
        return raster!.GetPixel(0, 0);
    }

    private static void SwapTagPayload(byte[] profile, string firstSignature, string secondSignature) {
        (int Offset, int Length) first = FindTag(profile, firstSignature);
        (int Offset, int Length) second = FindTag(profile, secondSignature);
        Assert.Equal(first.Length, second.Length);
        var temporary = new byte[first.Length];
        Buffer.BlockCopy(profile, first.Offset, temporary, 0, temporary.Length);
        Buffer.BlockCopy(profile, second.Offset, profile, first.Offset, temporary.Length);
        Buffer.BlockCopy(temporary, 0, profile, second.Offset, temporary.Length);
    }

    private static (int Offset, int Length) FindTag(byte[] profile, string signature) {
        uint target = ((uint)signature[0] << 24) | ((uint)signature[1] << 16) | ((uint)signature[2] << 8) | signature[3];
        int count = checked((int)ReadUInt32(profile, 128));
        for (int index = 0; index < count; index++) {
            int entry = 132 + index * 12;
            if (ReadUInt32(profile, entry) == target) {
                return (checked((int)ReadUInt32(profile, entry + 4)), checked((int)ReadUInt32(profile, entry + 8)));
            }
        }
        throw new InvalidOperationException("ICC tag was not found: " + signature + ".");
    }

    private static uint ReadUInt32(byte[] bytes, int offset) =>
        unchecked(((uint)bytes[offset] << 24) | ((uint)bytes[offset + 1] << 16) | ((uint)bytes[offset + 2] << 8) | bytes[offset + 3]);

    private static void WriteUInt32(byte[] bytes, int offset, uint value) {
        bytes[offset] = (byte)(value >> 24);
        bytes[offset + 1] = (byte)(value >> 16);
        bytes[offset + 2] = (byte)(value >> 8);
        bytes[offset + 3] = (byte)value;
    }

    private static void WriteS15Fixed16(byte[] bytes, int offset, double value) =>
        WriteUInt32(bytes, offset, unchecked((uint)(int)Math.Round(value * 65536D)));

    private static void WriteAscii(Stream stream, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }
}
