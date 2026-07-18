using OfficeIMO.Drawing;
using System;
using System.Collections.Generic;
using System.IO;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingRasterLimitTests {
    [Fact]
    public void RasterScaleLimiterHonorsCeilingRoundedPixelLimit() {
        OfficeRasterScaleLimit limit = OfficeRasterScaleLimiter.Resolve(3D, 3D, 1D, 2L);

        Assert.True(limit.WasLimited);
        Assert.Equal(1, limit.PixelWidth);
        Assert.Equal(1, limit.PixelHeight);
        Assert.True(limit.PixelCount <= 2L);
    }

    [Fact]
    public void RasterScaleLimiterHonorsPerDimensionLimit() {
        OfficeRasterScaleLimit limit = OfficeRasterScaleLimiter.Resolve(20_000D, 100D, 1D, 10_000_000L, 16_384);

        Assert.True(limit.WasLimited);
        Assert.Equal(16_384, limit.PixelWidth);
        Assert.True(limit.PixelHeight <= 16_384);
        Assert.True(limit.PixelCount <= 10_000_000L);
    }

    [Fact]
    public void JpegPixelLimitMatchesTheSharedEncoderByteCeiling() {
        Assert.Equal(33_554_432L, OfficeRasterImageEncoder.GetMaximumPixelCount(OfficeImageExportFormat.Jpeg));
        Assert.Equal(long.MaxValue, OfficeRasterImageEncoder.GetMaximumPixelCount(OfficeImageExportFormat.Png));
    }

    [Fact]
    public void FallbackCodecProducesVisibleContentAndStructuredDiagnostic() {
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var codec = new OfficeRasterImageFallbackCodec(diagnostics: diagnostics, source: "sample.svg");

        Assert.True(codec.TryDecode(new byte[] { 1, 2, 3 }, "image/svg+xml", out OfficeRasterImage? image));
        Assert.NotNull(image);
        Assert.NotEqual(OfficeColor.White, image!.GetPixel(0, 0));
        OfficeImageExportDiagnostic diagnostic = Assert.Single(diagnostics);
        Assert.Equal(OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback, diagnostic.Code);
        Assert.Equal("sample.svg", diagnostic.Source);
    }

    [Fact]
    public void PngDecoderRejectsOversizedHeaderBeforeInflatingImageData() {
        byte[] png = OfficePngWriter.Encode(new OfficeRasterImage(1, 1, OfficeColor.White));
        WriteBigEndian(png, 16, 100_000);
        WriteBigEndian(png, 20, 100_000);

        Assert.False(OfficePngReader.TryDecode(png, out OfficeRasterImage? image));
        Assert.Null(image);
    }

    [Fact]
    public void PngDecoderRejectsOverflowingChunkLengthBeforeAllocatingItsPayload() {
        byte[] png = {
            137, 80, 78, 71, 13, 10, 26, 10,
            0x7F, 0xFF, 0xFF, 0xFF,
            (byte)'P', (byte)'L', (byte)'T', (byte)'E',
            0, 0, 0, 0
        };

        Assert.False(OfficePngReader.TryDecode(png, out OfficeRasterImage? image));
        Assert.Null(image);
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void FallbackCodecHandlesMalformedAndIoFailuresFromApplicationCodecs(bool formatFailure) {
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        Exception failure = formatFailure
            ? new FormatException("Malformed external image.")
            : new IOException("External image source is unavailable.");
        var codec = new OfficeRasterImageFallbackCodec(new ThrowingCodec(failure), diagnostics, "external.img");

        Assert.True(codec.TryDecode(new byte[] { 1 }, "image/custom", out OfficeRasterImage? image));
        Assert.NotNull(image);
        OfficeImageExportDiagnostic diagnostic = Assert.Single(diagnostics);
        Assert.Equal(OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback, diagnostic.Code);
        Assert.Contains(failure.Message, diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void FallbackCodecReportsSuccessfulCallerCodecUse() {
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var codec = new OfficeRasterImageFallbackCodec(new SolidCodec(), diagnostics, "external.img");

        Assert.True(codec.TryDecode(new byte[] { 1 }, "image/custom", out OfficeRasterImage? image));
        Assert.NotNull(image);
        OfficeImageExportDiagnostic diagnostic = Assert.Single(diagnostics);
        Assert.Equal(OfficeImageExportDiagnosticCodes.SourceImageDecodedByCallerCodec, diagnostic.Code);
        Assert.Equal(OfficeImageExportDiagnosticSeverity.Info, diagnostic.Severity);
        Assert.Equal("external.img", diagnostic.Source);
    }

    [Fact]
    public void FallbackCodecHandlesOverflowFromApplicationCodec() {
        var diagnostics = new List<OfficeImageExportDiagnostic>();
        var codec = new OfficeRasterImageFallbackCodec(
            new ThrowingCodec(new OverflowException("External dimensions overflowed.")),
            diagnostics,
            "external.img");

        Assert.True(codec.TryDecode(new byte[] { 1 }, "image/custom", out OfficeRasterImage? image));
        Assert.NotNull(image);
        OfficeImageExportDiagnostic diagnostic = Assert.Single(diagnostics);
        Assert.Equal(OfficeImageExportDiagnosticCodes.SourceImageDecodeFallback, diagnostic.Code);
        Assert.Contains("overflowed", diagnostic.Message, StringComparison.Ordinal);
    }

    [Fact]
    public void RasterExportPlannerCombinesCallerAndEncoderLimitsBeforeAllocation() {
        var options = new OfficeImageExportOptions {
            Scale = 10D,
            MaximumRasterPixels = 1_000L
        };

        OfficeRasterExportPlan plan = OfficeRasterExportPlanner.Resolve(
            100D,
            100D,
            OfficeImageExportFormat.Webp,
            options,
            "test surface");

        Assert.True(plan.Limit.WasLimited);
        Assert.True(plan.Limit.PixelCount <= 1_000L);
        Assert.NotNull(plan.Diagnostic);
        Assert.Equal(OfficeImageExportDiagnosticCodes.RasterScaleReduced, plan.Diagnostic!.Code);
        Assert.Equal("test surface", plan.Diagnostic.Source);
    }

    [Fact]
    public void RasterExportPlannerCanRejectOversizedRequestsWithTypedEvidence() {
        var options = new OfficeImageExportOptions {
            Scale = 10D,
            MaximumRasterPixels = 1_000L,
            RasterOverflowBehavior = OfficeRasterOverflowBehavior.Throw
        };

        OfficeImageExportLimitException exception = Assert.Throws<OfficeImageExportLimitException>(() =>
            OfficeRasterExportPlanner.Resolve(
                100D,
                100D,
                OfficeImageExportFormat.Png,
                options));

        Assert.Equal(10D, exception.RequestedScale);
        Assert.Equal(1_000_000L, exception.RequestedPixels);
        Assert.Equal(1_000L, exception.MaximumPixels);
    }

    private static void WriteBigEndian(byte[] bytes, int offset, int value) {
        bytes[offset] = (byte)(value >> 24);
        bytes[offset + 1] = (byte)(value >> 16);
        bytes[offset + 2] = (byte)(value >> 8);
        bytes[offset + 3] = (byte)value;
    }

    private sealed class ThrowingCodec : IOfficeRasterImageCodec {
        private readonly Exception _exception;

        internal ThrowingCodec(Exception exception) => _exception = exception;

        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            image = null;
            throw _exception;
        }
    }

    private sealed class SolidCodec : IOfficeRasterImageCodec {
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            image = new OfficeRasterImage(1, 1, OfficeColor.Black);
            return true;
        }
    }
}
