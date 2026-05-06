using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public class DrawingTests {
    private static readonly byte[] OnePixelPng = {
        0x89, 0x50, 0x4E, 0x47, 0x0D, 0x0A, 0x1A, 0x0A,
        0x00, 0x00, 0x00, 0x0D, 0x49, 0x48, 0x44, 0x52,
        0x00, 0x00, 0x00, 0x01, 0x00, 0x00, 0x00, 0x01,
        0x08, 0x06, 0x00, 0x00, 0x00, 0x1F, 0x15, 0xC4,
        0x89, 0x00, 0x00, 0x00, 0x0A, 0x49, 0x44, 0x41,
        0x54, 0x78, 0x9C, 0x63, 0x00, 0x01, 0x00, 0x00,
        0x05, 0x00, 0x01, 0x0D, 0x0A, 0x2D, 0xB4, 0x00,
        0x00, 0x00, 0x00, 0x49, 0x45, 0x4E, 0x44, 0xAE,
        0x42, 0x60, 0x82
    };

    [Fact]
    public void OfficeColorParsesNamedAndHexValues() {
        Assert.Equal(OfficeColor.Red, OfficeColor.Parse("red"));

        var color = OfficeColor.Parse("#336699CC");
        Assert.Equal(0x33, color.R);
        Assert.Equal(0x66, color.G);
        Assert.Equal(0x99, color.B);
        Assert.Equal(0xCC, color.A);
        Assert.Equal("336699CC", color.ToHex());
        Assert.Equal("336699", color.ToRgbHex());
        Assert.Equal("CC336699", color.ToArgbHex());
    }

    [Fact]
    public void OfficeColorTransparentUsesZeroRgbWithZeroAlpha() {
        Assert.Equal(0, OfficeColor.Transparent.R);
        Assert.Equal(0, OfficeColor.Transparent.G);
        Assert.Equal(0, OfficeColor.Transparent.B);
        Assert.Equal(0, OfficeColor.Transparent.A);
    }

    [Fact]
    public void OfficeFontInfoStoresFamilySizeAndStyleWithoutFontEngineDependency() {
        var font = new OfficeFontInfo("Aptos", 12.5, OfficeFontStyle.Bold | OfficeFontStyle.Italic);

        Assert.Equal("Aptos", font.FamilyName);
        Assert.Equal(12.5, font.Size);
        Assert.True(font.IsBold);
        Assert.True(font.IsItalic);
        Assert.Equal("Aptos, 12.5pt, Bold, Italic", font.ToString());
    }

    [Fact]
    public void OfficeFontInfoProvidesImmutableCopyHelpers() {
        var font = OfficeFontInfo.Default
            .WithFamilyName("Arial")
            .WithSize(10)
            .WithStyle(OfficeFontStyle.Bold);

        Assert.Equal(new OfficeFontInfo("Arial", 10, OfficeFontStyle.Bold), font);
        Assert.NotEqual(OfficeFontInfo.Default, font);
    }

    [Theory]
    [InlineData("png", OfficeImageFormat.Png)]
    [InlineData(".png", OfficeImageFormat.Png)]
    [InlineData("photo.JPG", OfficeImageFormat.Jpeg)]
    [InlineData("diagram.svg", OfficeImageFormat.Svg)]
    [InlineData("legacy.emf", OfficeImageFormat.Emf)]
    public void OfficeImageReaderMapsFileNamesAndBareExtensions(string fileName, OfficeImageFormat expected) {
        Assert.Equal(expected, OfficeImageReader.FromExtension(fileName));
    }

    [Fact]
    public void OfficeImageReaderIdentifiesPngWithoutDecodingPixels() {
        var image = OfficeImageReader.Identify(OnePixelPng);

        Assert.Equal(OfficeImageFormat.Png, image.Format);
        Assert.Equal(1, image.Width);
        Assert.Equal(1, image.Height);
        Assert.Equal("image/png", image.MimeType);
    }

    [Fact]
    public void OfficeImageReaderIdentifiesIconDimensionsFromHeader() {
        var icon = new byte[22];
        icon[2] = 0x01;
        icon[4] = 0x01;
        icon[6] = 16;
        icon[7] = 32;

        var image = OfficeImageReader.Identify(icon);

        Assert.Equal(OfficeImageFormat.Icon, image.Format);
        Assert.Equal(16, image.Width);
        Assert.Equal(32, image.Height);
        Assert.Equal("image/x-icon", image.MimeType);
    }

    [Fact]
    public void OfficeImageReaderIdentifiesPcxDimensionsFromHeader() {
        var pcx = new byte[128];
        pcx[0] = 0x0A;
        pcx[1] = 0x05;
        pcx[2] = 0x01;
        pcx[3] = 0x08;
        pcx[8] = 99;
        pcx[10] = 49;
        pcx[12] = 96;
        pcx[14] = 96;

        var image = OfficeImageReader.Identify(pcx);

        Assert.Equal(OfficeImageFormat.Pcx, image.Format);
        Assert.Equal(100, image.Width);
        Assert.Equal(50, image.Height);
        Assert.Equal(96, image.DpiX);
        Assert.Equal(96, image.DpiY);
        Assert.Equal("image/x-pcx", image.MimeType);
    }

    [Fact]
    public void OfficeImageReaderIdentifiesEmfDimensionsFromHeader() {
        var emf = new byte[88];
        WriteInt32LittleEndian(emf, 0, 1);
        WriteInt32LittleEndian(emf, 4, 88);
        WriteInt32LittleEndian(emf, 16, 192);
        WriteInt32LittleEndian(emf, 20, 96);
        WriteInt32LittleEndian(emf, 32, 5080);
        WriteInt32LittleEndian(emf, 36, 2540);
        WriteInt32LittleEndian(emf, 40, 0x464D4520);
        WriteInt32LittleEndian(emf, 72, 1920);
        WriteInt32LittleEndian(emf, 76, 1080);
        WriteInt32LittleEndian(emf, 80, 508);
        WriteInt32LittleEndian(emf, 84, 286);

        var image = OfficeImageReader.Identify(emf);

        Assert.Equal(OfficeImageFormat.Emf, image.Format);
        Assert.Equal(192, image.Width);
        Assert.Equal(96, image.Height);
        Assert.Equal(96, Math.Round(image.DpiX));
        Assert.Equal(96, Math.Round(image.DpiY));
        Assert.Equal("image/x-emf", image.MimeType);
    }

    [Fact]
    public void OfficeImageReaderIdentifiesPlaceableWmfDimensionsFromHeader() {
        var wmf = new byte[22];
        WriteInt32LittleEndian(wmf, 0, unchecked((int)0x9AC6CDD7));
        WriteInt16LittleEndian(wmf, 10, 2880);
        WriteInt16LittleEndian(wmf, 12, 1440);
        WriteUInt16LittleEndian(wmf, 14, 1440);
        WritePlaceableWmfChecksum(wmf);

        var image = OfficeImageReader.Identify(wmf);

        Assert.Equal(OfficeImageFormat.Wmf, image.Format);
        Assert.Equal(192, image.Width);
        Assert.Equal(96, image.Height);
        Assert.Equal("image/x-wmf", image.MimeType);
    }

    [Fact]
    public void OfficeImageReaderRestoresSeekableStreamPosition() {
        var data = new byte[OnePixelPng.Length + 2];
        data[0] = 0xAA;
        data[1] = 0xBB;
        Array.Copy(OnePixelPng, 0, data, 2, OnePixelPng.Length);
        using var stream = new MemoryStream(data);
        stream.Position = 2;

        Assert.True(OfficeImageReader.TryIdentify(stream, null, out var image));

        Assert.Equal(2, stream.Position);
        Assert.Equal(OfficeImageFormat.Png, image.Format);
        Assert.Equal(1, image.Width);
        Assert.Equal(1, image.Height);
    }

    [Fact]
    public void OfficeImageReaderReadsSvgViewBoxDimensions() {
        var svg = System.Text.Encoding.UTF8.GetBytes("<svg xmlns=\"http://www.w3.org/2000/svg\" viewBox=\"0 0 320 180\"></svg>");

        Assert.True(OfficeImageReader.TryIdentify(svg, "chart.svg", out var image));

        Assert.Equal(OfficeImageFormat.Svg, image.Format);
        Assert.Equal(320, image.Width);
        Assert.Equal(180, image.Height);
        Assert.Equal("image/svg+xml", image.MimeType);
    }

    [Fact]
    public void OfficeImageReaderReadsSvgPhysicalUnitsAsPixels() {
        var svg = System.Text.Encoding.UTF8.GetBytes("<svg xmlns=\"http://www.w3.org/2000/svg\" width=\"1in\" height=\"2.54cm\"></svg>");

        Assert.True(OfficeImageReader.TryIdentify(svg, "units.svg", out var image));

        Assert.Equal(OfficeImageFormat.Svg, image.Format);
        Assert.Equal(96, image.Width);
        Assert.Equal(96, image.Height);
    }

    [Fact]
    public void OfficeImageReaderRejectsSvgWithDtdWhenNoExtensionFallbackExists() {
        var svg = System.Text.Encoding.UTF8.GetBytes("<?xml version=\"1.0\"?><!DOCTYPE svg [<!ENTITY xxe SYSTEM \"file:///c:/windows/win.ini\">]><svg xmlns=\"http://www.w3.org/2000/svg\" width=\"1\" height=\"1\">&xxe;</svg>");

        Assert.False(OfficeImageReader.TryIdentify(svg, null, out var image));

        Assert.Equal(OfficeImageFormat.Unknown, image.Format);
    }

    [Fact]
    public void OfficeImageReaderFallsBackToSvgExtensionWhenXmlCannotBeParsed() {
        var svg = System.Text.Encoding.UTF8.GetBytes("<?xml version=\"1.0\"?><!DOCTYPE svg [<!ENTITY xxe SYSTEM \"file:///c:/windows/win.ini\">]><svg xmlns=\"http://www.w3.org/2000/svg\" width=\"1\" height=\"1\">&xxe;</svg>");

        Assert.True(OfficeImageReader.TryIdentify(svg, "unsafe.svg", out var image));

        Assert.Equal(OfficeImageFormat.Svg, image.Format);
        Assert.Equal(0, image.Width);
        Assert.Equal(0, image.Height);
    }

    private static void WriteInt16LittleEndian(byte[] data, int offset, short value) {
        data[offset] = (byte)(value & 0xFF);
        data[offset + 1] = (byte)((value >> 8) & 0xFF);
    }

    private static void WriteUInt16LittleEndian(byte[] data, int offset, ushort value) {
        data[offset] = (byte)(value & 0xFF);
        data[offset + 1] = (byte)((value >> 8) & 0xFF);
    }

    private static void WriteInt32LittleEndian(byte[] data, int offset, int value) {
        data[offset] = (byte)(value & 0xFF);
        data[offset + 1] = (byte)((value >> 8) & 0xFF);
        data[offset + 2] = (byte)((value >> 16) & 0xFF);
        data[offset + 3] = (byte)((value >> 24) & 0xFF);
    }

    private static void WritePlaceableWmfChecksum(byte[] data) {
        ushort checksum = 0;
        for (int offset = 0; offset < 20; offset += 2) {
            checksum ^= (ushort)(data[offset] | (data[offset + 1] << 8));
        }

        WriteUInt16LittleEndian(data, 20, checksum);
    }
}
