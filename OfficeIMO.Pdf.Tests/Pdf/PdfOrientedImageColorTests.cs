using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed class PdfOrientedImageColorTests {
    [Fact]
    public void OpaqueMonochromeOrientationKeepsSamplesCompatibleWithItsRgbProfile() {
        byte[] jpeg = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "Images", "white-srgb-oriented.jpg"));
        Assert.True(PdfWriter.TryGetJpegComponentCount(jpeg, out int components));
        Assert.Equal(3, components);
        Assert.True(PdfDocument.TryPrepareImageBytes(jpeg, out byte[] normalized, out _, out _, out string? reason), reason);
        Assert.Equal(PdfIccProfiles.SrgbIec6196621, OfficeImageMetadata.Read(normalized).IccProfile);
        byte[] pdf = PdfDocument.Create().Image(jpeg, 20, 20).ToBytes();
        Assert.Equal(1, Assert.Single(PdfImageExtractor.ExtractImages(pdf)).Width);
    }

    [Theory]
    [InlineData(2, 2, 1)]
    [InlineData(6, 1, 2)]
    public void ColorManagedOrientationRetainsAnExplicitSrgbCalibration(int orientation, int width, int height) {
        var raster = new OfficeRasterImage(2, 1);
        raster.SetPixel(0, 0, OfficeColor.Red);
        raster.SetPixel(1, 0, OfficeColor.Blue);
        byte[] jpeg = OfficeJpegCodec.Encode(raster, new OfficeJpegEncodeOptions {
            Quality = 100,
            Subsampling = OfficeJpegSubsampling.Y444,
            Metadata = new OfficeJpegMetadata(icc: PdfIccProfiles.SrgbIec6196621, exif: new byte[] {
                (byte)'I', (byte)'I', 42, 0, 8, 0, 0, 0,
                1, 0, 0x12, 1, 3, 0, 1, 0, 0, 0,
                (byte)orientation, 0, 0, 0, 0, 0, 0, 0
            })
        });
        Assert.True(PdfDocument.TryPrepareImageBytes(jpeg, out byte[] normalized, out _, out bool transcoded, out string? reason), reason);
        Assert.True(transcoded);
        Assert.Equal(PdfIccProfiles.SrgbIec6196621, OfficeImageMetadata.Read(normalized).IccProfile);
        byte[] pdf = PdfDocument.Create().Image(jpeg, 20, 20).ToBytes();
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        PdfStream image = Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfStream>(),
            stream => stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Image");
        Assert.Equal(width, image.Dictionary.Get<PdfNumber>("Width")!.Value);
        Assert.Equal(height, image.Dictionary.Get<PdfNumber>("Height")!.Value);
        PdfArray space = Assert.IsType<PdfArray>(image.Dictionary.Items["ColorSpace"]);
        Assert.Equal("ICCBased", Assert.IsType<PdfName>(space.Items[0]).Name);
        PdfStream profile = Assert.IsType<PdfStream>(objects[Assert.IsType<PdfReference>(space.Items[1]).ObjectNumber].Value);
        Assert.Equal(3, profile.Dictionary.Get<PdfNumber>("N")!.Value);
        Assert.Equal(PdfIccProfiles.SrgbIec6196621, StreamDecoder.Decode(profile.Dictionary, profile.Data, objects));
        PdfExtractedImage extracted = Assert.Single(PdfImageExtractor.ExtractImages(pdf));
        Assert.True(OfficePngReader.TryDecode(extracted.Bytes, out OfficeRasterImage? decoded));
        OfficeColor first = decoded!.GetPixel(0, 0);
        OfficeColor last = decoded.GetPixel(width - 1, height - 1);
        Assert.True(orientation == 2 ? first.B > first.R && last.R > last.B : first.R > first.B && last.B > last.R);
    }
}
