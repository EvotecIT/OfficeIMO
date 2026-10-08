using OfficeIMO.Drawing;
using OfficeIMO.Epub;
using OfficeIMO.Html;
using System.IO.Compression;
using System.Security.Cryptography;
using System.Text.Json;
using Xunit;

namespace OfficeIMO.Workflows.Tests;

public sealed class BookKdpDeliveryTests {
    [Theory]
    [InlineData(OfficeImageExportFormat.Jpeg, "listing-cover.jpeg")]
    [InlineData(OfficeImageExportFormat.Tiff, "listing-cover.tiff")]
    public void PacketPreservesCoverAndPublicationWithExplicitAcceptanceGaps(OfficeImageExportFormat format, string name) {
        byte[] cover = Cover(format);
        var project = BookProject.Create("Published title");
        project.CreateRevision("Private history");
        byte[] before = project.ToProjectBytes();
        var options = new EpubWriteOptions { ModifiedAt = new DateTimeOffset(2026, 10, 5, 0, 0, 0, TimeSpan.Zero) };
        var ordinary = Entries(project.ToDeliveryBytes(options));
        byte[] packet = project.ToKdpDeliveryBytes(cover, options);
        var entries = Entries(packet);
        Assert.Equal(new[] { "manifest.json", name, "package.opf", "publication.epub" }.OrderBy(x => x), entries.Keys.OrderBy(x => x));
        Assert.Equal(cover, entries[name]);
        Assert.Equal(ordinary["publication.epub"], entries["publication.epub"]);
        Assert.Equal(ordinary["package.opf"], entries["package.opf"]);
        using var json = JsonDocument.Parse(entries["manifest.json"]);
        var root = json.RootElement;
        Assert.Equal("OfficeIMO.KdpDelivery", root.GetProperty("Format").GetString());
        Assert.Equal("not-performed", root.GetProperty("KindlePreviewer").GetString());
        Assert.Equal("not-checked", root.GetProperty("Cover").GetProperty("ColorModeAndSeparation").GetString());
        Assert.Equal("passed", root.GetProperty("Cover").GetProperty("ManagedPixelDecode").GetString());
        Assert.Equal(format == OfficeImageExportFormat.Jpeg ? "three-component-jpeg" : "rgb-tiff",
            root.GetProperty("Cover").GetProperty("EncodedColorStructure").GetString());
        Assert.Equal("not-performed", root.GetProperty("Delivery").GetProperty("RetailerAcceptance").GetString());
        foreach (var file in root.GetProperty("Delivery").GetProperty("Files").EnumerateArray()) {
            byte[] bytes = entries[file.GetProperty("Name").GetString()!];
            Assert.Equal(bytes.LongLength, file.GetProperty("Bytes").GetInt64());
            Assert.Equal(Convert.ToHexString(SHA256.HashData(bytes)), file.GetProperty("Sha256").GetString());
        }
        Assert.Equal(packet, project.ToKdpDeliveryBytes(cover, options));
        Assert.Equal(before, project.ToProjectBytes());
    }

    [Theory]
    [InlineData(624, 1000)]
    [InlineData(625, 999)]
    [InlineData(10001, 1000)]
    [InlineData(625, 10001)]
    [InlineData(5000, 5000)]
    public void RejectsDimensionAndLocalPixelLimits(int width, int height) {
        // Patch TIFF dimensions without allocating an oversized alleged raster.
        byte[] cover = Cover(OfficeImageExportFormat.Tiff);
        SetTiffDimension(cover, 256, width);
        SetTiffDimension(cover, 257, height);
        Assert.Throws<InvalidDataException>(() => BookProject.Create("Book").ToKdpDeliveryBytes(cover));
    }

    [Fact]
    public void RejectsUnsupportedFormatAndHeaderOnlyImage() {
        var project = BookProject.Create("Book");
        byte[] png = Cover(OfficeImageExportFormat.Png);
        Assert.Throws<InvalidDataException>(() => project.ToKdpDeliveryBytes(png));
        byte[] jpeg = Cover(OfficeImageExportFormat.Jpeg);
        Assert.Throws<InvalidDataException>(() => project.ToKdpDeliveryBytes(jpeg[..(jpeg.Length / 2)]));
        Assert.Throws<InvalidDataException>(() => project.ToKdpDeliveryBytes([]));
        Assert.Throws<ArgumentNullException>(() => project.ToKdpDeliveryBytes(null!));
    }

    [Theory]
    [InlineData(5)]
    [InlineData(6)]
    [InlineData(7)]
    [InlineData(8)]
    public void RotatedPortraitUsesDisplayDimensionsAndPreservesSource(int orientation) {
        byte[] cover = OrientedJpeg(1000, 625, orientation);
        var entries = Entries(BookProject.Create("Book").ToKdpDeliveryBytes(cover));
        Assert.Equal(cover, entries["listing-cover.jpeg"]);
        using var json = JsonDocument.Parse(entries["manifest.json"]);
        var checks = json.RootElement.GetProperty("Cover");
        Assert.Equal(625, checks.GetProperty("Width").GetInt32());
        Assert.Equal(1000, checks.GetProperty("Height").GetInt32());
        Assert.Equal("decoded-display-after-embedded-orientation", checks.GetProperty("DimensionBasis").GetString());
        // 1000 / 625 meets the recommended ratio, despite the stored landscape axes.
        Assert.Single(checks.GetProperty("Recommendations").EnumerateArray());
    }

    [Theory]
    [InlineData(5)]
    [InlineData(6)]
    [InlineData(7)]
    [InlineData(8)]
    public void RotatedLandscapeCannotPassUsingUnrotatedPortraitDimensions(int orientation) {
        byte[] cover = OrientedJpeg(625, 1000, orientation);
        var error = Assert.Throws<InvalidDataException>(() => BookProject.Create("Book").ToKdpDeliveryBytes(cover));
        Assert.Contains("must display", error.Message);
    }

    private static byte[] OrientedJpeg(int width, int height, int orientation) {
        byte[] jpeg = Cover(OfficeImageExportFormat.Jpeg, width, height);
        // Standard APP1 EXIF: little-endian TIFF with one SHORT Orientation tag.
        byte[] app1 = [0xff, 0xe1, 0, 34, 69, 120, 105, 102, 0, 0,
            73, 73, 42, 0, 8, 0, 0, 0, 1, 0, 18, 1, 3, 0, 1, 0, 0, 0,
            (byte)orientation, 0, 0, 0, 0, 0, 0, 0];
        return [.. jpeg[..2], .. app1, .. jpeg[2..]];
    }

    [Fact]
    public void DecodableGrayscaleJpegRequiresAnRgbSourceCover() {
        byte[] cover = OfficeRasterImageEncoder.Encode(
            new OfficeRasterImage(625, 1000, OfficeColor.White), OfficeImageExportFormat.Jpeg);
        Assert.Equal(1, OfficeImageReader.Identify(cover).JpegComponentCount);
        Assert.True(OfficeRasterImageDecoder.TryDecode(cover, out _));
        var error = Assert.Throws<InvalidDataException>(() => BookProject.Create("Book").ToKdpDeliveryBytes(cover));
        Assert.Contains("three-component JPEG or explicitly RGB TIFF", error.Message);
    }

    [Fact]
    public void DecodableCmykTiffCannotPassThroughTheRgbaDecoder() {
        byte[] cover = Cover(OfficeImageExportFormat.Tiff);
        int offset = BitConverter.ToInt32(cover, 4);
        int count = BitConverter.ToUInt16(cover, offset);
        bool changed = false, removedAlpha = false;
        for (int i = 0; i < count; i++) {
            int entry = offset + 2 + 12 * i;
            ushort tag = BitConverter.ToUInt16(cover, entry);
            if (tag == 262) {
                BitConverter.GetBytes((ushort)5).CopyTo(cover, entry + 8);
                changed = true;
            }
            if (tag == 338) {
                // Reinterpret the four encoded channels as CMYK, not RGB plus alpha.
                Assert.Equal(count - 1, i);
                BitConverter.GetBytes((ushort)(count - 1)).CopyTo(cover, offset);
                Array.Clear(cover, entry, 4); // terminal next-IFD pointer after the shortened table
                removedAlpha = true;
            }
        }
        Assert.True(changed && removedAlpha);
        Assert.True(OfficeRasterImageDecoder.TryDecode(cover, out var image));
        Assert.Equal(625, image!.Width);
        Assert.Equal(1000, image.Height);
        var error = Assert.Throws<InvalidDataException>(() => BookProject.Create("Book").ToKdpDeliveryBytes(cover));
        Assert.Contains("three-component JPEG or explicitly RGB TIFF", error.Message);
    }

    [Fact]
    public void RejectsMultiPageCoverWithoutSelectingOnlyFirstPage() {
        var page = new OfficeRasterImage(625, 1000, OfficeColor.White);
        byte[] cover = OfficeTiffCodec.EncodePages([page, page]);
        Assert.Throws<InvalidDataException>(() => BookProject.Create("Book").ToKdpDeliveryBytes(cover));
    }

    [Fact]
    public void RejectsCoverAtExclusiveByteLimit() {
        byte[] cover = new byte[BookProject.MaximumKdpCoverBytes + 1];
        Assert.Throws<InvalidDataException>(() => BookProject.Create("Book").ToKdpDeliveryBytes(cover));
    }

    [Fact]
    public void RecommendationsDoNotRejectOtherwiseSupportedCover() {
        byte[] cover = Cover(OfficeImageExportFormat.Tiff, 1000, 1000);
        using var json = JsonDocument.Parse(Entries(BookProject.Create("Book").ToKdpDeliveryBytes(cover))["manifest.json"]);
        Assert.Equal(2, json.RootElement.GetProperty("Cover").GetProperty("Recommendations").GetArrayLength());
    }

    [Fact]
    public void ImportReviewAndWriterBoundsStillApplyWithoutMutation() {
        var project = BookProject.FromImport(EpubManuscript.ImportHtml(HtmlConversionDocument.Parse(
            "<title>Book</title><h1>Chapter</h1><script>ignored()</script>")));
        byte[] cover = Cover(OfficeImageExportFormat.Tiff);
        Assert.Throws<InvalidOperationException>(() => project.ToKdpDeliveryBytes(cover));
        project.AcknowledgeImportLoss();
        byte[] before = project.ToProjectBytes();
        var options = new EpubWriteOptions { MaxOutputBytes = 256L * 1024 * 1024 };
        byte[] packet = project.ToKdpDeliveryBytes(cover, options);
        Assert.Throws<InvalidDataException>(() => project.ToKdpDeliveryBytes(cover, options, packet.Length - 1));
        Assert.Throws<InvalidDataException>(() => project.ToKdpDeliveryBytes(cover, new EpubWriteOptions { MaxOutputBytes = 16 }));
        Assert.Equal(256L * 1024 * 1024, options.MaxOutputBytes);
        Assert.Throws<ArgumentOutOfRangeException>(() => project.ToKdpDeliveryBytes(cover, maximumOutputBytes: 0));
        Assert.Throws<ArgumentOutOfRangeException>(() => project.ToKdpDeliveryBytes(cover, maximumOutputBytes: BookProject.MaximumKdpDeliveryBytes + 1));
        using var cancellation = new CancellationTokenSource(); cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => project.ToKdpDeliveryBytes(cover, cancellationToken: cancellation.Token));
        Assert.Equal(before, project.ToProjectBytes());
    }

    private static byte[] Cover(OfficeImageExportFormat format, int width = 625, int height = 1000) =>
        OfficeRasterImageEncoder.Encode(new OfficeRasterImage(width, height, OfficeColor.CornflowerBlue), format);

    private static Dictionary<string, byte[]> Entries(byte[] bytes) {
        using var stream = new MemoryStream(bytes, false);
        using var zip = new ZipArchive(stream, ZipArchiveMode.Read);
        return zip.Entries.ToDictionary(entry => entry.FullName, entry => {
            using var input = entry.Open(); using var output = new MemoryStream();
            input.CopyTo(output); return output.ToArray();
        });
    }

    private static void SetTiffDimension(byte[] bytes, ushort tag, int value) {
        Assert.Equal((byte)'I', bytes[0]);
        int offset = BitConverter.ToInt32(bytes, 4);
        int count = BitConverter.ToUInt16(bytes, offset);
        for (int i = 0; i < count; i++) {
            int entry = offset + 2 + i * 12;
            if (BitConverter.ToUInt16(bytes, entry) != tag) continue;
            Assert.Equal(4, BitConverter.ToUInt16(bytes, entry + 2));
            BitConverter.GetBytes(value).CopyTo(bytes, entry + 8);
            return;
        }
        throw new InvalidOperationException("TIFF dimension tag not found.");
    }
}
