using System.Globalization;
using System.Text;
using OfficeIMO.Core.Internal;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfIccColorRenderingTests {
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
    public void RenderDiagnostics_RejectsDctThatRequiresExternalIccConversion() {
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

        Assert.Contains(
            page.GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
    }

    [Fact]
    public void DctPassThroughRejectsNonIdentityDecodeForRgbIccFallback() {
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

        Assert.Contains(
            document.Pages[0].GetRenderCapabilityDiagnostics(),
            diagnostic => diagnostic.Code == PdfRenderCapabilities.ColorSpaceId);
        Assert.False(Assert.Single(PdfImageExtractor.ExtractImages(pdf)).IsImageFile);
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
        string colorSpaceResources = "/CsIcc [/ICCBased 5 0 R]") {
        string content = "/" + colorSpaceName + " cs\n" + colorOperation + "\n40 80 70 40 re\nf";
        byte[] contentBytes = Encoding.ASCII.GetBytes(content);
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.4\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /ColorSpace << " + colorSpaceResources + " >> >> /Contents 4 0 R >>\nendobj\n");
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
        int bitsPerComponent = 8) {
        byte[] contentBytes = Encoding.ASCII.GetBytes("q\n40 0 0 40 40 80 cm\n/Im1 Do\nQ");
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.4\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /BitsPerComponent " + bitsPerComponent.ToString(CultureInfo.InvariantCulture) + " /ColorSpace " + imageColorSpace + " " + imageEntries + " /Length " + imageSamples.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
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

    private static byte[] Compress(byte[] bytes) => OfficeZlibCodec.Compress(bytes);

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

    private static void WriteAscii(Stream stream, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }
}
