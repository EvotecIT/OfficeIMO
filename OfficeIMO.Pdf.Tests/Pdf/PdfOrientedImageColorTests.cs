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
        Assert.True(PdfDocument.TryPrepareImageBytes(jpeg, out byte[] prepared, out _, out bool transcoded, out string? reason), reason);
        Assert.False(transcoded);
        Assert.Equal(jpeg, prepared);
        Assert.Equal(PdfIccProfiles.SrgbIec6196621, OfficeImageMetadata.Read(prepared).IccProfile);
        byte[] pdf = PdfDocument.Create().Image(jpeg, 20, 20).ToBytes();
        PdfStream image = Assert.Single(PdfSyntax.ParseObjects(pdf).Map.Values.Select(item => item.Value).OfType<PdfStream>(),
            stream => stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Image");
        Assert.Equal(2D, image.Dictionary.Get<PdfNumber>("Width")!.Value);
        Assert.Equal(1D, image.Dictionary.Get<PdfNumber>("Height")!.Value);
        Assert.Equal(jpeg, image.Data);
    }

    [Theory]
    [InlineData(1, 2, 1)]
    [InlineData(2, 2, 1)]
    [InlineData(3, 2, 1)]
    [InlineData(4, 2, 1)]
    [InlineData(5, 1, 2)]
    [InlineData(6, 1, 2)]
    [InlineData(7, 1, 2)]
    [InlineData(8, 1, 2)]
    public void ColorManagedOrientationPreservesDctProfileAndRawDimensions(int orientation, int width, int height) {
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
        Assert.True(PdfDocument.TryPrepareImageBytes(jpeg, out byte[] prepared, out OfficeImageInfo? info, out bool transcoded, out string? reason), reason);
        Assert.False(transcoded);
        Assert.Equal(jpeg, prepared);
        Assert.Equal(width, info!.Width);
        Assert.Equal(height, info.Height);
        Assert.Equal(PdfIccProfiles.SrgbIec6196621, OfficeImageMetadata.Read(prepared).IccProfile);
        byte[] pdf = PdfDocument.Create().Image(jpeg, width * 20, height * 20).ToBytes();
        var (objects, _) = PdfSyntax.ParseObjects(pdf);
        PdfStream image = Assert.Single(objects.Values.Select(item => item.Value).OfType<PdfStream>(),
            stream => stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Image");
        Assert.Equal(2D, image.Dictionary.Get<PdfNumber>("Width")!.Value);
        Assert.Equal(1D, image.Dictionary.Get<PdfNumber>("Height")!.Value);
        Assert.Equal("DCTDecode", image.Dictionary.Get<PdfName>("Filter")?.Name);
        Assert.Equal(jpeg, image.Data);
        PdfArray space = Assert.IsType<PdfArray>(image.Dictionary.Items["ColorSpace"]);
        Assert.Equal("ICCBased", Assert.IsType<PdfName>(space.Items[0]).Name);
        PdfStream profile = Assert.IsType<PdfStream>(objects[Assert.IsType<PdfReference>(space.Items[1]).ObjectNumber].Value);
        Assert.Equal(3, profile.Dictionary.Get<PdfNumber>("N")!.Value);
        Assert.Equal(PdfIccProfiles.SrgbIec6196621, StreamDecoder.Decode(profile.Dictionary, profile.Data, objects));
        Assert.Equal(jpeg, image.Data);
        PdfImagePlacement placement = Assert.Single(PdfImageExtractor.ExtractImagePlacements(pdf));
        Assert.Equal(width * 20D, placement.Width, 3);
        Assert.Equal(height * 20D, placement.Height, 3);
        // Mirror-only orientations must reverse handedness; quarter-turns exchange the paint axes.
        double determinant = placement.A * placement.D - placement.B * placement.C;
        Assert.Equal(orientation is 2 or 4 or 5 or 7 ? -1 : 1, Math.Sign(determinant));
        Assert.Equal(orientation >= 5, Math.Abs(placement.A) < 0.001 && Math.Abs(placement.D) < 0.001);
    }
}
