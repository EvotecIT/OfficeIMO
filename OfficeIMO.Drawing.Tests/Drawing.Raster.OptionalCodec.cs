using System;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRasterOptionalCodecTests {
    // LibTIFF 4.7.2 independently encoded and decoded 16x16 RGB, LZMA compression.
    private const string IndependentLzmaTiff = "SUkqAFAAAAD9N3pYWgAAAP8S2UECAQMBACEBFnkgxO7gAv8ADV0Af4A8Fz4mR/wBtzwgAAAAAAAAASGABgAAAADtKJuoAAr8AgAAAAAAWVoKAAABAwABAAAAEAAAAAEBAwABAAAAEAAAAAIBAwADAAAAzgAAAAMBAwABAAAAbYgAAAYBAwABAAAAAgAAABEBBAABAAAACAAAABUBAwABAAAAAwAAABYBAwABAAAAEAAAABcBBAABAAAASAAAABwBAwABAAAAAQAAAAAAAAAIAAgACAA=";
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

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ProgressiveArithmeticJpegUsesManagedPixels(bool supplyCallerCodec) {
        // libjpeg-turbo 3.2.0 cjpeg -arithmetic -progressive; constant 16x12 RGB (30,80,120).
        byte[] bytes = Convert.FromBase64String("/9j/4AAQSkZJRgABAQAAAQABAAD/2wBDAAgGBgcGBQgHBwcJCQgKDBQNDAsLDBkSEw8UHRofHh0aHBwgJC4nICIsIxwcKDcpLDAxNDQ0Hyc5PTgyPC4zNDL/2wBDAQkJCQwLDBgNDRgyIRwhMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjIyMjL/ygARCAAMABADASIAAhEBAxEB/8wABgAQARD/2gAMAwEAAhADEAAAAf8AKhWW5P/MAAQQBf/aAAgBAQABBQLA/8wABBEF/9oACAEDAQE/AcD/zAAEEQX/2gAIAQIBAT8BwP/MAAQQBf/aAAgBAQAGPwLA/8wABBAF/9oACAEBAAE/IcD/2gAMAwEAAgADAAAAEGD/zAAEEQX/2gAIAQMBAT8QwP/MAAQRBf/aAAgBAgEBPxDA/8wABBAF/9oACAEBAAE/EMD/2Q==");
        Assert.True(OfficeImageReader.TryIdentifyByContent(bytes, null, out var source));
        Assert.Equal(16, source.Width);
        var drawing = new OfficeDrawing(16, 12).AddImage(bytes, "image/jpeg",
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, 16, 12)));
        var codec = supplyCallerCodec ? new IncorrectJpegCodec() : null;
        var result = drawing.ExportImage(OfficeImageExportFormat.Png,
            new OfficeImageExportOptions { BackgroundColor = OfficeColor.Transparent, ImageCodec = codec });
        Assert.True(OfficePngReader.TryDecode(result.Bytes, out var image));
        Assert.True(image!.GetPixel(8, 6).A > 0);
        Assert.DoesNotContain(result.Diagnostics, x => x.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
        Assert.InRange(image.GetPixel(8, 6).B, 118, 122);
        if (codec != null) Assert.Equal(0, codec.Calls);
    }

    [Fact]
    public void InspectedCallerSuccessRetainsProvenanceDiagnostic() {
        byte[] bytes = Convert.FromBase64String(IndependentAnimatedWebp);
        var drawing = new OfficeDrawing(16, 16).AddImageWithInterpolation(bytes, "image/webp",
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, 16, 16)), interpolate: false);
        foreach (var format in new[] { OfficeImageExportFormat.Png, OfficeImageExportFormat.Svg }) {
            var result = drawing.ExportImage(format, new OfficeImageExportOptions { ImageCodec = new Codec() });
            Assert.Contains(result.Diagnostics, d => d.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodedByCallerCodec);
            Assert.DoesNotContain(result.Diagnostics, d => d.Code == OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback);
            Assert.Contains(result.Diagnostics, d => d.Code == OfficeImageExportDiagnosticCodes.SourceImageStaticFrameSelected);
        }
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void SharedDecoderNeverAcceptsFallbackPlaceholders(bool nested) {
        // Independently encoded 32x32 lossless animation (Pillow/libwebp).
        byte[] bytes = Convert.FromBase64String("UklGRoQAAABXRUJQVlA4WAoAAAACAAAAHwAAHwAAQU5JTQYAAAAAAAAAAABBTk1GKAAAAAAAAAAAAB8AAB8AAGQAAAJWUDhMDwAAAC8fwAcABxD9j/4HIqL/AQBBTk1GKAAAAAAAAAAAAB8AAB8AAGQAAABWUDhMDwAAAC8fwAcABxDR//4HIqL/AQA=");
        Assert.True(OfficeRasterContainerInspector.TryInspect(bytes, out var container));
        Assert.Equal(32, container!.CanvasWidth);
        IOfficeRasterImageCodec codec = new OfficeRasterImageFallbackCodec();
        if (nested) codec = new OfficeRasterImageFallbackCodec(codec);
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, new OfficeRasterDecodeOptions { ImageCodec = codec }, out var image, out _));
        Assert.Null(image);
        Assert.False(OfficeImagePngConverter.TryConvertToPng(bytes, new OfficeRasterDecodeOptions { ImageCodec = codec }, out _, out _));
    }

    [Theory]
    [InlineData(16, true)]
    [InlineData(1, false)]
    public void InspectedUnsupportedTiffUsesValidatedFirstPageCallerPixels(int size, bool succeeds) {
        // Valid independent LZMA TIFF remains outside managed compression support.
        byte[] bytes = Convert.FromBase64String(IndependentLzmaTiff);
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, out _));
        var codec = new TiffCodec(size);
        Assert.Equal(succeeds, OfficeRasterImageDecoder.TryDecode(bytes,
            new OfficeRasterDecodeOptions { ImageCodec = codec }, out var image, out var info));
        Assert.Equal(1, codec.Calls);
        Assert.Equal(succeeds, info.UsedCallerCodec);
        if (succeeds) Assert.Equal(OfficeColor.Red, image!.GetPixel(8, 8));
        else Assert.Null(image);
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes,
            new OfficeRasterDecodeOptions { ImageCodec = codec, MaximumDecodedPixels = 1 }, out _, out _));
        Assert.Equal(1, codec.Calls);
    }

    private sealed class TiffCodec(int size) : IOfficeRasterImageCodec {
        internal int Calls;
        public bool TryDecode(byte[] bytes, string? contentType, out OfficeRasterImage? image) {
            Calls++;
            Assert.Equal("image/tiff", contentType);
            image = new OfficeRasterImage(size, size, OfficeColor.Red);
            return true;
        }
    }

    [Fact]
    public void OptionalCodecCloneAndOutputMustFitRetainedBudgetBeforeCallback() {
        byte[] bytes = Convert.FromBase64String(IndependentLzmaTiff);
        Array.Resize(ref bytes, 128 * 1024); // TIFF permits unreferenced trailing data.
        var codec = new TiffCodec(16);
        // The inspector's 64 KiB allowance fits, but the 128 KiB provider input clone does not.
        long retained = OfficeRasterGuards.MaximumDecodedBytes - bytes.Length - 65536 - 1024;
        var options = new OfficeRasterDecodeOptions { ImageCodec = codec, RetainedManagedBytes = retained };
        Assert.True(OfficeRasterContainerInspector.TryInspectForDecode(bytes, options, out _, out _));
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes,
            new OfficeRasterDecodeOptions { ImageCodec = codec, RetainedManagedBytes = retained }, out var image, out _));
        Assert.Null(image);
        Assert.Equal(0, codec.Calls);
    }

    private sealed class IncorrectJpegCodec : IOfficeRasterImageCodec {
        public int Calls;
        public bool TryDecode(byte[] bytes, string? contentType, out OfficeRasterImage? image) {
            Calls++;
            image = new OfficeRasterImage(1, 1, OfficeColor.Red);
            return true;
        }
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
