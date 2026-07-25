using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingFontContainerTests {
    [Fact]
    public void OfficeFontContainerDecoder_RoundTripsCompressedWoffIntoReusableOpenType() {
        byte[] source = ManagedTextShapingTestAssets.CreateFont('A', 0x1F600);
        int headOffset = FindTableOffset(source, "head");
        WriteUInt32(source, headOffset + 8, 0xBADB455D);
        byte[] woff = ManagedTextShapingTestAssets.CreateWoff(source);

        bool decoded = OfficeFontContainerDecoder.TryDecodeToOpenType(
            woff,
            out byte[] openType,
            out OfficeFontContainerFormat format,
            out string? error);

        Assert.True(decoded, error);
        Assert.Equal(OfficeFontContainerFormat.Woff, format);
        Assert.NotSame(source, openType);
        OfficeTrueTypeFont font = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(openType));
        Assert.True(font.HasGlyphs("A" + char.ConvertFromUtf32(0x1F600)));
        Assert.Equal(0xB1B0AFBA, CalculateChecksum(openType));
        Assert.NotEqual(0U, ReadUInt32(openType, FindTableOffset(openType, "head") + 8));
        var fonts = new OfficeFontFaceCollection().Add("WOFF Demo", woff);
        Assert.Single(fonts.Faces);
        Assert.Equal(OfficeFontContainerFormat.OpenType, OfficeFontContainerDecoder.Detect(fonts.Faces[0].Data));
    }

    [Fact]
    public void OfficeFontContainerDecoder_RejectsMalformedOrOversizedContainersWithoutPartialOutput() {
        byte[] source = ManagedTextShapingTestAssets.CreateFont('A');
        byte[] woff = ManagedTextShapingTestAssets.CreateWoff(source);
        int firstTableOffset = (woff[48] << 24) | (woff[49] << 16) | (woff[50] << 8) | woff[51];
        woff[firstTableOffset] ^= 0x7F;

        Assert.False(OfficeFontContainerDecoder.TryDecodeToOpenType(
            woff,
            out byte[] decoded,
            out OfficeFontContainerFormat format,
            out string? error));
        Assert.Equal(OfficeFontContainerFormat.Woff, format);
        Assert.Empty(decoded);
        Assert.False(string.IsNullOrWhiteSpace(error));

        byte[] valid = ManagedTextShapingTestAssets.CreateWoff(source);
        Assert.False(OfficeFontContainerDecoder.TryDecodeToOpenType(
            valid,
            source.Length - 1,
            out decoded,
            out format,
            out error));
        Assert.Empty(decoded);
        Assert.Contains("limit", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void OfficeFontContainerDecoder_DetectsWoff2AndReportsItsCurrentBoundary() {
        byte[] woff2 = { 0x77, 0x4F, 0x46, 0x32, 0, 0, 0, 0 };

        Assert.Equal(OfficeFontContainerFormat.Woff2, OfficeFontContainerDecoder.Detect(woff2));
        Assert.False(OfficeFontContainerDecoder.TryDecodeToOpenType(
            woff2,
            out byte[] decoded,
            out OfficeFontContainerFormat format,
            out string? error));
        Assert.Equal(OfficeFontContainerFormat.Woff2, format);
        Assert.Empty(decoded);
        Assert.Contains("WOFF 2", error, StringComparison.Ordinal);
    }

    private static int FindTableOffset(byte[] font, string tag) {
        int tableCount = (font[4] << 8) | font[5];
        for (int index = 0; index < tableCount; index++) {
            int record = 12 + index * 16;
            if (font[record] == tag[0] && font[record + 1] == tag[1]
                && font[record + 2] == tag[2] && font[record + 3] == tag[3]) {
                return checked((int)ReadUInt32(font, record + 8));
            }
        }
        throw new InvalidOperationException("The test font has no " + tag + " table.");
    }

    private static uint CalculateChecksum(byte[] data) {
        uint checksum = 0;
        for (int offset = 0; offset < data.Length; offset += 4) {
            uint value = (uint)data[offset] << 24;
            if (offset + 1 < data.Length) value |= (uint)data[offset + 1] << 16;
            if (offset + 2 < data.Length) value |= (uint)data[offset + 2] << 8;
            if (offset + 3 < data.Length) value |= data[offset + 3];
            checksum = unchecked(checksum + value);
        }
        return checksum;
    }

    private static uint ReadUInt32(byte[] data, int offset) =>
        ((uint)data[offset] << 24)
        | ((uint)data[offset + 1] << 16)
        | ((uint)data[offset + 2] << 8)
        | data[offset + 3];

    private static void WriteUInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }
}
