using System;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRasterOptionalCodecTests {
    // Independently encoded two-frame 16x16 animation (Pillow 12.3.0, libwebp 1.6.0).
    private const string IndependentAnimatedWebp = "UklGRogAAABXRUJQVlA4WAoAAAACAAAADwAADwAAQU5JTQYAAAAAAAAAAABBTk1GKgAAAAAAAAAAAA8AAA8AAGQAAAJWUDhMEQAAAC8PwAMAB1CoohSv/4GI6H8AAEFOTUYqAAAAAAAAAAAADwAADwAAZAAAAFZQOEwRAAAALw/AAwAHUKjiFaX/gYjofwAA";

    [Fact]
    public void OptionalCodecNormalizesInspectedPayloadWithoutMutatingInput() {
        byte[] bytes = Convert.FromBase64String(IndependentAnimatedWebp);
        byte[] original = (byte[])bytes.Clone();
        var codec = new Codec { MutateInput = true };
        Assert.True(OfficeImagePngConverter.TryConvertToPng(bytes,
            new OfficeRasterDecodeOptions { ImageCodec = codec }, out byte[] png, out var info));
        Assert.Equal(1, codec.Calls);
        Assert.Equal(original, bytes);
        Assert.True(info.Succeeded);
        Assert.Equal(OfficeImageFormat.Webp, info.Format);
        Assert.True(OfficePngReader.TryDecode(png, out var image));
        Assert.Equal(OfficeColor.Red, image!.GetPixel(8, 8));

        var svgCodec = new Codec { MutateInput = true };
        var drawing = new OfficeDrawing(16, 16).AddImageWithInterpolation(bytes, "image/webp",
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, 16, 16)), interpolate: false);
        string svg = OfficeDrawingSvgExporter.ToSvg(drawing, 1D, OfficeSvgSizeUnit.Pixel, imageCodec: svgCodec);
        Assert.Contains("fill=\"#FF0000\"", svg);
        Assert.Equal(1, svgCodec.Calls);
        Assert.Equal(original, bytes);
    }

    [Theory]
    [InlineData(255, 128)]
    [InlineData(256, 1)]
    public void ResourceLimitsRefuseBeforeInvokingOptionalCodec(long pixels, int encodedBytes) {
        var codec = new Codec();
        Assert.False(OfficeRasterImageDecoder.TryDecode(Convert.FromBase64String(IndependentAnimatedWebp),
            new OfficeRasterDecodeOptions { ImageCodec = codec, MaximumDecodedPixels = pixels, MaximumEncodedBytes = encodedBytes },
            out _, out _));
        Assert.Equal(0, codec.Calls);
    }

    [Fact]
    public void OptionalCodecCancellationIsObservedAfterCallback() {
        using var cts = new CancellationTokenSource();
        var codec = new Codec { AfterDecode = cts.Cancel };
        Assert.Throws<OperationCanceledException>(() => OfficeRasterImageDecoder.TryDecode(
            Convert.FromBase64String(IndependentAnimatedWebp),
            new OfficeRasterDecodeOptions { ImageCodec = codec, CancellationToken = cts.Token }, out _, out _));
        Assert.Equal(1, codec.Calls);
    }

    [Fact]
    public void OptionalCodecCannotReturnDimensionsDifferentFromInspectedContainer() {
        var codec = new Codec { Size = 1 };
        Assert.False(OfficeRasterImageDecoder.TryDecode(Convert.FromBase64String(IndependentAnimatedWebp),
            new OfficeRasterDecodeOptions { ImageCodec = codec }, out var image, out _));
        Assert.Null(image);
        Assert.Equal(1, codec.Calls);

        var svgCodec = new Codec { Size = 1 };
        var drawing = new OfficeDrawing(16, 16).AddImageWithInterpolation(
            Convert.FromBase64String(IndependentAnimatedWebp), "image/webp",
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, 16, 16)), interpolate: false);
        Assert.Throws<InvalidOperationException>(() => OfficeDrawingSvgExporter.ToSvg(
            drawing, 1D, OfficeSvgSizeUnit.Pixel, imageCodec: svgCodec));
        Assert.Equal(1, svgCodec.Calls);
    }

    [Fact]
    public void MalformedContainerIsRejectedBeforeOptionalCodec() {
        var codec = new Codec();
        byte[] bytes = Convert.FromBase64String(IndependentAnimatedWebp);
        Array.Resize(ref bytes, bytes.Length - 1);
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes,
            new OfficeRasterDecodeOptions { ImageCodec = codec }, out _, out _));
        Assert.Equal(0, codec.Calls);
    }

    [Fact]
    public void OptionalCodecFormatFailureReturnsFailedDecode() {
        var codec = new Codec { AfterDecode = () => throw new FormatException("Unsupported coding feature") };
        Assert.False(OfficeRasterImageDecoder.TryDecode(Convert.FromBase64String(IndependentAnimatedWebp),
            new OfficeRasterDecodeOptions { ImageCodec = codec }, out var image, out _));
        Assert.Null(image);
        Assert.Equal(1, codec.Calls);
    }

    [Fact]
    public void BuiltInDecodeDoesNotInvokeOptionalCodec() {
        var codec = new Codec();
        byte[] bytes = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.Blue));
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes,
            new OfficeRasterDecodeOptions { ImageCodec = codec }, out var image, out _));
        Assert.Equal(OfficeColor.Blue, image!.GetPixel(0, 0));
        Assert.Equal(0, codec.Calls);
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void PublicDrawingExportRetainsVisiblePlaceholderOutsideSourceDimensionValidation(bool invalidCodecDimensions) {
        byte[] bytes = Convert.FromBase64String(IndependentAnimatedWebp);
        var codec = invalidCodecDimensions ? new Codec { Size = 1 } : null;
        var drawing = new OfficeDrawing(16, 16).AddImage(bytes, "image/webp",
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, 16, 16)));
        var result = drawing.ExportImage(OfficeImageExportFormat.Png,
            new OfficeImageExportOptions { ImageCodec = codec, BackgroundColor = OfficeColor.Transparent });
        Assert.True(OfficePngReader.TryDecode(result.Bytes, out var image));
        Assert.True(image!.GetPixel(image.Width / 2, image.Height / 2).A > 0);
        Assert.Contains(result.Diagnostics, x => x.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback
            && x.LossKind == OfficeConversionLossKind.Omission);
        if (codec != null) Assert.Equal(1, codec.Calls);

        var nearest = new OfficeDrawing(16, 16).AddImageWithInterpolation(bytes, "image/webp",
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, 16, 16)), interpolate: false);
        var svgCodec = invalidCodecDimensions ? new Codec { Size = 1 } : null;
        var svgResult = nearest.ExportImage(OfficeImageExportFormat.Svg,
            new OfficeImageExportOptions { ImageCodec = svgCodec, BackgroundColor = OfficeColor.Transparent });
        Assert.Contains(svgResult.Diagnostics, x => x.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback
            && x.LossKind == OfficeConversionLossKind.Omission);
        Assert.DoesNotContain(svgResult.Diagnostics, x => x.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodedByCallerCodec);
        Assert.Contains("#BE4646", System.Text.Encoding.UTF8.GetString(svgResult.Bytes));
        if (svgCodec != null) Assert.Equal(1, svgCodec.Calls);
    }

    private sealed class Codec : IOfficeRasterImageCodec {
        internal int Calls;
        internal bool MutateInput;
        internal int Size = 16;
        internal Action? AfterDecode;
        public bool TryDecode(byte[] bytes, string? contentType, out OfficeRasterImage? image) {
            Calls++;
            Assert.Equal("image/webp", contentType);
            if (MutateInput) Array.Clear(bytes, 0, bytes.Length);
            image = new OfficeRasterImage(Size, Size, OfficeColor.Red);
            AfterDecode?.Invoke();
            return true;
        }
    }
}
