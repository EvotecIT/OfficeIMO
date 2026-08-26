using System;
using System.IO;
using System.Threading;
using OfficeIMO.Drawing;
using OfficeIMO.TestAssets;
using Xunit;

namespace OfficeIMO.Drawing.Tests;

public sealed class DrawingWoff2FontTests {
    [Fact]
    public void Woff1RetainsContainerIdentityAndUsesFirstPartyCffProgram() {
        byte[] openType = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "SourceSansPro-Regular.otf"));
        byte[] woff = ManagedTextShapingTestAssets.CreateWoff(openType);
        OfficeFontFace face = Assert.Single(new OfficeFontFaceCollection().Add("Source Sans WOFF", woff).Faces);

        Assert.Equal(OfficeFontContainerFormat.Woff, face.ContainerFormat);
        Assert.True(face.CanEmbedAsStaticPdfFont);
        Assert.True(face.Program.IsOpenTypeCff);
        Assert.True(face.Program.HasGlyphs("OfficeIMO ffi"));
    }

#if NET8_0_OR_GREATER
    [Fact]
    public void Woff2ContainerDecodesToUsableTrueTypeProgram() {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "OpenSans-Regular.woff2"));

        bool decoded = OfficeFontContainerDecoder.TryDecodeToOpenType(
            source,
            16 * 1024 * 1024,
            out byte[] openType,
            out OfficeFontContainerFormat format,
            out string? error);

        Assert.True(decoded, error);
        Assert.Equal(OfficeFontContainerFormat.Woff2, format);
        Assert.Equal(OfficeFontContainerFormat.OpenType, OfficeFontContainerDecoder.Detect(openType));
        Assert.True(openType.Length > source.Length);
        OfficeTrueTypeFont font = Assert.IsType<OfficeTrueTypeFont>(OfficeTrueTypeFont.TryLoad(openType));
        Assert.True(font.HasGlyphs("OfficeIMO ffi 123"));
        Assert.True(font.Measure("OfficeIMO ffi 123", 24D) > 100D);
        Assert.NotEmpty(font.GetTextContours("OfficeIMO", 0D, 0D, 24D));
    }

    [Fact]
    public void Woff2CollectionMeasuresRasterizesAndHonorsOutlineBudgets() {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "OpenSans-Regular.woff2"));
        var fonts = new OfficeFontFaceCollection();
        Assert.True(fonts.TryAddBounded(
            "Open Sans WOFF2",
            source,
            OfficeFontStyle.Regular,
            OfficeFontUnicodeRangeSet.All,
            16 * 1024 * 1024,
            out int decodedBytes,
            out string? error), error);
        OfficeFontFace face = Assert.Single(fonts.Faces);
        Assert.Equal(OfficeFontContainerFormat.Woff2, face.ContainerFormat);
        Assert.True(face.CanEmbedAsStaticPdfFont);
        Assert.True(decodedBytes > source.Length);

        var image = new OfficeRasterImage(360, 80, OfficeColor.White);
        var canvas = new OfficeRasterCanvas(image, fonts: fonts);
        canvas.DrawText("OfficeIMO ffi 123", 0D, 0D, image.Width, image.Height, OfficeColor.Black, 36D, fontFamily: "Open Sans WOFF2");
        Assert.Contains(image.GetPixels(), value => value < 250);

        IOfficeBoundedFontProgram bounded = Assert.IsAssignableFrom<IOfficeBoundedFontProgram>(face.Program);
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.ThrowsAny<OperationCanceledException>(() => bounded.GetTextContoursBounded(
            "OfficeIMO", 0D, 0D, 24D, 10_000, cancellation.Token));
        Assert.Throws<InvalidOperationException>(() => bounded.GetTextContoursBounded(
            "OfficeIMO", 0D, 0D, 24D, 1, CancellationToken.None));
    }

    [Fact]
    public void Woff2DecoderRejectsTruncationAndDecodedSizeBudget() {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "OpenSans-Regular.woff2"));
        byte[] truncated = new byte[32];
        Buffer.BlockCopy(source, 0, truncated, 0, truncated.Length);

        Assert.False(OfficeFontContainerDecoder.TryDecodeToOpenType(
            truncated,
            16 * 1024 * 1024,
            out _,
            out OfficeFontContainerFormat format,
            out string? truncatedError));
        Assert.Equal(OfficeFontContainerFormat.Woff2, format);
        Assert.Contains("truncated", truncatedError, StringComparison.OrdinalIgnoreCase);

        Assert.False(OfficeFontContainerDecoder.TryDecodeToOpenType(
            source,
            maximumDecodedBytes: source.Length,
            out _,
            out _,
            out string? budgetError));
        Assert.Contains("limit", budgetError, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Woff2DecoderRejectsOversizedDirectoryOutputBeforeDecompression() {
        byte[] oversizedDirectory = CreateWoff2WithOversizedDeclaredGlyf();

        Assert.False(OfficeFontContainerDecoder.TryDecodeToOpenType(
            oversizedDirectory,
            maximumDecodedBytes: 1_024,
            out _,
            out OfficeFontContainerFormat format,
            out string? error));
        Assert.Equal(OfficeFontContainerFormat.Woff2, format);
        Assert.Contains("decoded WOFF 2 font exceeds", error, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Woff2DecoderTreatsDeclaredSfntSizeAsReferenceOnly() {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "OpenSans-Regular.woff2"));
        foreach (uint declaredSize in new[] { 0U, 1U, uint.MaxValue }) {
            byte[] referenceOnly = (byte[])source.Clone();
            WriteBigEndianUInt32(referenceOnly, 16, declaredSize);

            Assert.True(OfficeFontContainerDecoder.TryDecodeToOpenType(
                referenceOnly,
                16 * 1024 * 1024,
                out _,
                out OfficeFontContainerFormat format,
                out string? error), error);
            Assert.Equal(OfficeFontContainerFormat.Woff2, format);
        }
    }

    [Fact]
    public void Woff2DecoderAcceptsReservedHeaderValueButRejectsUnknownTableTransform() {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "OpenSans-Regular.woff2"));
        byte[] reservedHeader = (byte[])source.Clone();
        reservedHeader[14] = 0x12;
        reservedHeader[15] = 0x34;

        Assert.True(OfficeFontContainerDecoder.TryDecodeToOpenType(
            reservedHeader,
            16 * 1024 * 1024,
            out _,
            out _,
            out string? reservedError), reservedError);

        byte[] unknownTransform = (byte[])source.Clone();
        unknownTransform[48] = (byte)((unknownTransform[48] & 0x3F) | 0x40);
        Assert.False(OfficeFontContainerDecoder.TryDecodeToOpenType(
            unknownTransform,
            16 * 1024 * 1024,
            out _,
            out _,
            out string? transformError));
        Assert.Contains("transform version", transformError, StringComparison.OrdinalIgnoreCase);
    }

    [Fact]
    public void Woff2DecoderRejectsExtraneousTrailingData() {
        byte[] source = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "OpenSans-Regular.woff2"));
        byte[] withTrailingData = new byte[source.Length + 4];
        Buffer.BlockCopy(source, 0, withTrailingData, 0, source.Length);
        WriteBigEndianUInt32(withTrailingData, 8, checked((uint)withTrailingData.Length));

        Assert.False(OfficeFontContainerDecoder.TryDecodeToOpenType(
            withTrailingData,
            16 * 1024 * 1024,
            out _,
            out _,
            out string? error));
        Assert.Contains("extraneous", error, StringComparison.OrdinalIgnoreCase);
    }

    private static void WriteBigEndianUInt32(byte[] data, int offset, uint value) {
        data[offset] = (byte)(value >> 24);
        data[offset + 1] = (byte)(value >> 16);
        data[offset + 2] = (byte)(value >> 8);
        data[offset + 3] = (byte)value;
    }

    private static byte[] CreateWoff2WithOversizedDeclaredGlyf() {
        using var output = new MemoryStream();
        output.Write(new byte[48], 0, 48);
        output.WriteByte(10); // Known-tag index for transformed glyf.
        WriteBase128(output, 2_000_000);
        WriteBase128(output, 1);
        output.WriteByte(0);
        while ((output.Length & 3) != 0) output.WriteByte(0);
        byte[] bytes = output.ToArray();
        WriteBigEndianUInt32(bytes, 0, 0x774F4632);
        WriteBigEndianUInt32(bytes, 4, 0x00010000);
        WriteBigEndianUInt32(bytes, 8, checked((uint)bytes.Length));
        bytes[12] = 0;
        bytes[13] = 1;
        WriteBigEndianUInt32(bytes, 20, 1);
        return bytes;
    }

    private static void WriteBase128(Stream output, uint value) {
        var encoded = new byte[5];
        int offset = encoded.Length;
        do {
            encoded[--offset] = (byte)(value & 0x7F);
            value >>= 7;
        } while (value != 0);
        for (int index = offset; index < encoded.Length; index++) {
            byte current = encoded[index];
            if (index < encoded.Length - 1) current |= 0x80;
            output.WriteByte(current);
        }
    }
#endif
}
