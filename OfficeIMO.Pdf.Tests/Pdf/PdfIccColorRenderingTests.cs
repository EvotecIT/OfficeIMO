using System.Globalization;
using System.IO.Compression;
using System.Text;
using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public class PdfIccColorRenderingTests {
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
        const string alternate = "[/Lab << /WhitePoint [0.9505 1 1.089] /Range [-100 100 -100 100] >>]";
        byte[] pdf = BuildIccContentPdf(unsupportedProfile, "/N 3 /Alternate " + alternate, "50 40 -40 scn");

        OfficeDrawing drawing = PdfPageImageRenderer.RenderPage(pdf);

        OfficeColor expected = OfficeColorSpaceConverter.FromLab(50, 40, -40);
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

    private static byte[] BuildIccContentPdf(byte[] profile, string profileEntries, string colorOperation) {
        string content = "/CsIcc cs\n" + colorOperation + "\n40 80 70 40 re\nf";
        byte[] contentBytes = Encoding.ASCII.GetBytes(content);
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.4\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /ColorSpace << /CsIcc [/ICCBased 5 0 R] >> >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< " + profileEntries + " /Length " + profile.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(profile, 0, profile.Length);
        WriteAscii(output, "\nendstream\nendobj\ntrailer\n<< /Root 1 0 R >>\n%%EOF\n");
        return output.ToArray();
    }

    private static byte[] BuildIccImagePdf(
        byte[] profile,
        byte[] imageSamples,
        string profileEntries,
        string imageEntries = "",
        byte? softMaskSample = null,
        string imageColorSpace = "[/ICCBased 6 0 R]",
        string extraObjects = "") {
        byte[] contentBytes = Encoding.ASCII.GetBytes("q\n40 0 0 40 40 80 cm\n/Im1 Do\nQ");
        using var output = new MemoryStream();
        WriteAscii(output, "%PDF-1.4\n");
        WriteAscii(output, "1 0 obj\n<< /Type /Catalog /Pages 2 0 R >>\nendobj\n");
        WriteAscii(output, "2 0 obj\n<< /Type /Pages /Count 1 /Kids [3 0 R] /MediaBox [0 0 240 200] >>\nendobj\n");
        WriteAscii(output, "3 0 obj\n<< /Type /Page /Parent 2 0 R /Resources << /XObject << /Im1 5 0 R >> >> /Contents 4 0 R >>\nendobj\n");
        WriteAscii(output, "4 0 obj\n<< /Length " + contentBytes.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
        output.Write(contentBytes, 0, contentBytes.Length);
        WriteAscii(output, "\nendstream\nendobj\n");
        WriteAscii(output, "5 0 obj\n<< /Type /XObject /Subtype /Image /Width 1 /Height 1 /BitsPerComponent 8 /ColorSpace " + imageColorSpace + " " + imageEntries + " /Length " + imageSamples.Length.ToString(CultureInfo.InvariantCulture) + " >>\nstream\n");
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

    private static byte[] Compress(byte[] bytes) {
        using var output = new MemoryStream();
        using (var compression = new ZLibStream(output, CompressionLevel.SmallestSize, leaveOpen: true)) {
            compression.Write(bytes, 0, bytes.Length);
        }
        return output.ToArray();
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

    private static void WriteAscii(Stream stream, string value) {
        byte[] bytes = Encoding.ASCII.GetBytes(value);
        stream.Write(bytes, 0, bytes.Length);
    }
}
