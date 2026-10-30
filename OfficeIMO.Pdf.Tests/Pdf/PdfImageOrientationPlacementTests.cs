using OfficeIMO.Drawing;
using OfficeIMO.Pdf;
using OfficeIMO.Pdf.Filters;
using Xunit;

namespace OfficeIMO.Tests.Pdf;

public sealed partial class PdfImageOrientationPlacementTests {
    public static IEnumerable<object[]> Placements() {
        foreach (string path in new[] { "flow", "canvas", "drawing", "floating", "table", "inline", "header", "footer", "stamp", "watermark", "page-watermark" }) {
            for (int orientation = 1; orientation <= 8; orientation++) yield return new object[] { path, orientation };
        }
    }

    [Theory]
    [MemberData(nameof(Placements))]
    public void PublicPlacementKeepsOriginalDctAndAppliesExifToThePaintMatrix(string path, int orientation) {
        byte[] jpeg = CreateJpeg(orientation);
        byte[] pdf = Compose(path, jpeg);
        PdfStream resource = Assert.Single(Images(PdfSyntax.ParseObjects(pdf).Map));
        Assert.Equal(jpeg, resource.Data);
        Assert.Equal("DCTDecode", resource.Dictionary.Get<PdfName>("Filter")?.Name);
        Assert.Equal(3D, resource.Dictionary.Get<PdfNumber>("Width")!.Value);
        Assert.Equal(2D, resource.Dictionary.Get<PdfNumber>("Height")!.Value);

        PdfImagePlacement placement = Assert.Single(PdfImageExtractor.ExtractImagePlacements(pdf));
        OfficeTransform expected = ExpectedOrientation(orientation).Then(OfficeTransform.Scale(60, 40));
        AssertLinear(expected, placement);
        Assert.Equal(60D, placement.Width, 3);
        Assert.Equal(40D, placement.Height, 3);
        Assert.Equal(placement.X + expected.OffsetX, placement.E, 3);
        Assert.Equal(placement.Y + expected.OffsetY, placement.F, 3);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    [InlineData(6)]
    [InlineData(7)]
    [InlineData(8)]
    public void PreparationSeparatesOrientedLayoutSizeFromUnchangedJpegResource(int orientation) {
        byte[] jpeg = CreateJpeg(orientation);
        Assert.True(PdfDocument.TryPrepareImageBytes(jpeg, out byte[] prepared, out OfficeImageInfo? info,
            out bool transcoded, out string? reason), reason);
        Assert.False(transcoded);
        Assert.Equal(jpeg, prepared);
        Assert.Equal(orientation >= 5 ? 2 : 3, info!.Width);
        Assert.Equal(orientation >= 5 ? 3 : 2, info.Height);
        Assert.Equal(orientation >= 5 ? 72D : 144D, info.DpiX, 3);
        Assert.Equal(orientation >= 5 ? 144D : 72D, info.DpiY, 3);
    }

    public static IEnumerable<object[]> CroppedMirrors() {
        for (int orientation = 1; orientation <= 8; orientation++) {
            foreach (string path in new[] { "flow", "canvas", "drawing" }) {
                yield return new object[] { path, orientation, false, false };
                if (path != "flow") {
                    yield return new object[] { path, orientation, true, false };
                    yield return new object[] { path, orientation, false, true };
                    yield return new object[] { path, orientation, true, true };
                }
            }
        }
    }

    [Theory]
    [MemberData(nameof(CroppedMirrors))]
    public void CropAndCallerMirrorsActOnTheOrientedImage(string path, int orientation, bool horizontal, bool vertical) {
        byte[] jpeg = CreateJpeg(orientation);
        var style = new PdfImageStyle { SourceCrop = new PdfImageSourceCrop(0.25, 0.125, 0, 0.25) };
        byte[] pdf = PdfDocument.Create(builder => builder.Content(content => {
            if (path == "flow") content.Image(jpeg, 60, 40, style: style);
            else if (path == "canvas") content.Canvas(canvas => canvas.Image(jpeg, 10, 20, 60, 40,
                style: style, horizontalFlip: horizontal, verticalFlip: vertical));
            else {
                var projection = new OfficeImageProjection(new OfficeImagePlacement(10, 20, 60, 40),
                    new OfficeImageSourceCrop(0.25, 0.125, 0, 0.25), flipHorizontal: horizontal, flipVertical: vertical);
                content.Drawing(new OfficeDrawing(100, 100).AddImage(jpeg, "image/jpeg", projection));
            }
        }), Options()).ToBytes();

        PdfImagePlacement placement = Assert.Single(PdfImageExtractor.ExtractImagePlacements(pdf));
        // In normalized bottom-left coordinates: orientation, crop expansion, caller mirror, target box.
        OfficeTransform expected = ExpectedOrientation(orientation)
            .Then(new OfficeTransform(1 / 0.75, 0, 0, 1 / 0.625, -0.25 / 0.75, -0.25 / 0.625))
            .Then(new OfficeTransform(horizontal ? -1 : 1, 0, 0, vertical ? -1 : 1, horizontal ? 1 : 0, vertical ? 1 : 0))
            .Then(OfficeTransform.Scale(60, 40));
        AssertLinear(expected, placement);
        if (path == "canvas") {
            Assert.Equal(10D + expected.OffsetX, placement.E, 3);
            Assert.Equal(240D + expected.OffsetY, placement.F, 3); // pageHeight - top - height
        }
        Assert.NotNull(placement.Clip);
        Assert.Equal(jpeg, Assert.Single(Images(PdfSyntax.ParseObjects(pdf).Map)).Data);
    }

    [Theory]
    [InlineData(1)]
    [InlineData(2)]
    [InlineData(3)]
    [InlineData(4)]
    [InlineData(5)]
    [InlineData(6)]
    [InlineData(7)]
    [InlineData(8)]
    public void Png16ExifRetainsSourceProfileColorPrecisionAndSoftMask(int orientation) {
        byte[] png = CreatePng16(orientation);
        Assert.True(PdfDocument.TryPrepareImageBytes(png, out byte[] prepared, out OfficeImageInfo? info,
            out bool transcoded, out string? reason), reason);
        Assert.False(transcoded);
        Assert.Equal(png, prepared);
        Assert.Equal(orientation >= 5 ? 1 : 2, info!.Width);
        Assert.Equal(orientation >= 5 ? 2 : 1, info.Height);

        foreach (string path in new[] { "flow", "stamp" }) {
            byte[] pdf = Compose(path, png);
            var objects = PdfSyntax.ParseObjects(pdf).Map;
            PdfStream image = Images(objects).Single(stream => stream.Dictionary.Items.ContainsKey("SMask"));
            Assert.Equal(2D, image.Dictionary.Get<PdfNumber>("Width")!.Value);
            Assert.Equal(1D, image.Dictionary.Get<PdfNumber>("Height")!.Value);
            Assert.Equal(16D, image.Dictionary.Get<PdfNumber>("BitsPerComponent")!.Value);
            Assert.Equal(16D, image.Dictionary.Get<PdfDictionary>("DecodeParms")!.Get<PdfNumber>("BitsPerComponent")!.Value);
            Assert.Equal(new byte[] { 0, 0x12, 0x34, 0x80, 0x01, 0xff, 0xfe, 0xab, 0xcd, 0x01, 0x23, 0x45, 0x67 }, FlateDecoder.Decode(image.Data));
            PdfStream mask = Assert.IsType<PdfStream>(objects[Assert.IsType<PdfReference>(image.Dictionary.Items["SMask"]).ObjectNumber].Value);
            Assert.Equal(16D, mask.Dictionary.Get<PdfNumber>("BitsPerComponent")!.Value);
            Assert.Equal(2D, mask.Dictionary.Get<PdfNumber>("Width")!.Value);
            Assert.Equal(1D, mask.Dictionary.Get<PdfNumber>("Height")!.Value);
            Assert.Equal(new byte[] { 0, 0x40, 0x01, 0xc0, 0x23 }, FlateDecoder.Decode(mask.Data));
            PdfArray space = Assert.IsType<PdfArray>(image.Dictionary.Items["ColorSpace"]);
            Assert.Equal("ICCBased", Assert.IsType<PdfName>(space.Items[0]).Name);
            PdfStream profile = Assert.IsType<PdfStream>(objects[Assert.IsType<PdfReference>(space.Items[1]).ObjectNumber].Value);
            Assert.Equal(PdfIccProfiles.SrgbIec6196621, StreamDecoder.Decode(profile.Dictionary, profile.Data, objects));
            AssertLinear(ExpectedOrientation(orientation).Then(OfficeTransform.Scale(60, 40)),
                Assert.Single(PdfImageExtractor.ExtractImagePlacements(pdf)));
        }
    }

    [Theory]
    [InlineData(2)]
    [InlineData(6)]
    [InlineData(8)]
    public void OpaquePng16OrientationDoesNotRewriteItsIdat(int orientation) {
        byte[] png = CreatePng16(orientation, alpha: false);
        byte[] pdf = Compose("flow", png);
        PdfStream image = Assert.Single(Images(PdfSyntax.ParseObjects(pdf).Map));
        byte[] scanline = { 0, 0x12, 0x34, 0x80, 0x01, 0xff, 0xfe, 0xab, 0xcd, 0x01, 0x23, 0x45, 0x67 };
        Assert.Equal("FlateDecode", image.Dictionary.Get<PdfName>("Filter")?.Name);
        Assert.Equal(PdfPngTestImages.BuildStoredZlib(scanline), image.Data);
        Assert.Equal(16D, image.Dictionary.Get<PdfNumber>("BitsPerComponent")!.Value);
        Assert.False(image.Dictionary.Items.ContainsKey("SMask"));
    }

    private static byte[] Compose(string path, byte[] source) {
        byte[] pdf = PdfDocument.Create(builder => builder.Page(page => {
            if (path == "header") page.Header(header => header.Image(source, 60, 40));
            if (path == "footer") page.Footer(footer => footer.Image(source, 60, 40));
            if (path == "page-watermark") page.ImageWatermark(source, 60, 40, opacity: 1, rotationAngle: 0);
            page.Content(content => {
                switch (path) {
                    case "flow": content.Image(source, 60, 40); break;
                    case "canvas": content.Canvas(canvas => canvas.Image(source, 10, 20, 60, 40)); break;
                    case "drawing":
                        content.Drawing(new OfficeDrawing(100, 100).AddImage(source, "image/jpeg",
                            new OfficeImageProjection(new OfficeImagePlacement(10, 20, 60, 40))));
                        break;
                    case "table":
                    case "floating":
                        var style = PdfTableStyles.Minimal();
                        style.HeaderRowCount = 0;
                        style.CellPaddingX = style.CellPaddingY = global::OfficeIMO.OfficeLength.Points(0);
                        if (path == "floating") {
                            style.Position = new PdfTablePosition();
                            style.ColumnWidths = new List<global::OfficeIMO.OfficeLength?> { global::OfficeIMO.OfficeLength.Points(100) };
                        }
                        content.Table(new[] { new[] { PdfTableCell.WithImages(string.Empty,
                            new[] { new PdfTableCellImage(source, 60, 40) }) } }, style: style);
                        break;
                    case "inline": content.Paragraph(paragraph => paragraph.InlineImage(source, 60, 40, fit: OfficeImageFit.Stretch)); break;
                    default: content.Paragraph(paragraph => paragraph.Text("Image target")); break;
                }
            });
        }), Options()).ToBytes();
        var stamp = new PdfImageStampOptions { X = 10, Y = 20, Width = 60, Height = 40 };
        return path == "stamp" ? PdfStamper.StampImage(pdf, source, stamp)
            : path == "watermark" ? PdfStamper.WatermarkImage(pdf, source, stamp) : pdf;
    }

    private static PdfCreateOptions Options() => new() {
        FileVersion = PdfFileVersion.Pdf17, PageWidth = global::OfficeIMO.OfficeLength.Points(300), PageHeight = global::OfficeIMO.OfficeLength.Points(300),
        MarginLeft = global::OfficeIMO.OfficeLength.Points(60), MarginRight = global::OfficeIMO.OfficeLength.Points(60), MarginTop = global::OfficeIMO.OfficeLength.Points(60), MarginBottom = global::OfficeIMO.OfficeLength.Points(60)
    };

    private static void AssertLinear(OfficeTransform expected, PdfImagePlacement placement) {
        Assert.Equal(expected.M11, placement.A, 3);
        Assert.Equal(expected.M12, placement.B, 3);
        Assert.Equal(expected.M21, placement.C, 3);
        Assert.Equal(expected.M22, placement.D, 3);
    }

    // Independent EXIF corner mapping, with PDF's bottom-left origin. Do not derive this table from the writer.
    private static OfficeTransform ExpectedOrientation(int value) => value switch {
        1 => new(1, 0, 0, 1, 0, 0),
        2 => new(-1, 0, 0, 1, 1, 0),
        3 => new(-1, 0, 0, -1, 1, 1),
        4 => new(1, 0, 0, -1, 0, 1),
        5 => new(0, -1, -1, 0, 1, 1),
        6 => new(0, -1, 1, 0, 0, 1),
        7 => new(0, 1, 1, 0, 0, 0),
        8 => new(0, 1, -1, 0, 1, 0),
        _ => throw new ArgumentOutOfRangeException(nameof(value))
    };

    private static byte[] CreateJpeg(int orientation) {
        var raster = new OfficeRasterImage(3, 2, OfficeColor.Red);
        raster.SetPixel(2, 0, OfficeColor.Blue);
        raster.SetPixel(0, 1, OfficeColor.Green);
        return OfficeJpegCodec.Encode(raster, new OfficeJpegEncodeOptions {
            Quality = 95, Subsampling = OfficeJpegSubsampling.Y444, DpiX = 144, DpiY = 72,
            Metadata = new OfficeJpegMetadata(exif: Exif(orientation))
        });
    }

    private static byte[] CreatePng16(int orientation, bool alpha = true) {
        using var output = new MemoryStream();
        byte[] signature = { 137, 80, 78, 71, 13, 10, 26, 10 };
        output.Write(signature, 0, signature.Length);
        PdfPngTestImages.WritePngChunk(output, "IHDR", new byte[] { 0, 0, 0, 2, 0, 0, 0, 1, 16, (byte)(alpha ? 6 : 2), 0, 0, 0 });
        PdfPngTestImages.WritePngChunk(output, "eXIf", Exif(orientation));
        byte[] profile = PdfPngTestImages.BuildStoredZlib(PdfIccProfiles.SrgbIec6196621);
        byte[] iccp = new byte[profile.Length + 5];
        Encoding.ASCII.GetBytes("RGB").CopyTo(iccp, 0);
        profile.CopyTo(iccp, 5);
        PdfPngTestImages.WritePngChunk(output, "iCCP", iccp);
        byte[] samples = alpha ? new byte[] {
            0, 0x12, 0x34, 0x80, 0x01, 0xff, 0xfe, 0x40, 0x01, 0xab, 0xcd, 0x01, 0x23, 0x45, 0x67, 0xc0, 0x23
        } : new byte[] { 0, 0x12, 0x34, 0x80, 0x01, 0xff, 0xfe, 0xab, 0xcd, 0x01, 0x23, 0x45, 0x67 };
        PdfPngTestImages.WritePngChunk(output, "IDAT", PdfPngTestImages.BuildStoredZlib(samples));
        PdfPngTestImages.WritePngChunk(output, "IEND", Array.Empty<byte>());
        return output.ToArray();
    }

    private static byte[] Exif(int orientation) => new byte[] {
        (byte)'I', (byte)'I', 42, 0, 8, 0, 0, 0,
        1, 0, 0x12, 1, 3, 0, 1, 0, 0, 0, (byte)orientation, 0, 0, 0, 0, 0, 0, 0
    };

    private static IEnumerable<PdfStream> Images(Dictionary<int, PdfIndirectObject> objects) =>
        objects.Values.Select(item => item.Value).OfType<PdfStream>()
            .Where(stream => stream.Dictionary.Get<PdfName>("Subtype")?.Name == "Image");
}
