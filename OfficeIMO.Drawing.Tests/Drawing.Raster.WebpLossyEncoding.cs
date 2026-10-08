using System;
using System.IO;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingWebpLossyEncodingTests {
    [Theory]
    [InlineData(1, 1)]
    [InlineData(1, 17)]
    [InlineData(17, 1)]
    [InlineData(33, 25)]
    public void LossyWebpPreservesVisibleDimensionsAndEveryAlphaSample(int width, int height) {
        OfficeRasterImage source = CreateImage(width, height, withAlpha: true);
        byte[] original = source.GetPixels();
        byte[] bytes = OfficeWebpCodec.Encode(source, Lossy(85));
        Assert.True(OfficeWebpCodec.TryDecode(bytes, out OfficeRasterImage? result));
        Assert.NotNull(result);
        Assert.Equal(width, result!.Width);
        Assert.Equal(height, result.Height);
        byte[] decoded = result.GetPixels();
        for (int i = 3; i < original.Length; i += 4) Assert.Equal(original[i], decoded[i]);
        Assert.Equal(original, source.GetPixels());
        Assert.Equal(bytes, OfficeWebpCodec.Encode(source, Lossy(85)));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void QualityChangesColorErrorAndEncodedSizeWithoutSwitchingToLossless(bool withAlpha) {
        OfficeRasterImage source = CreateImage(129, 97, withAlpha);
        byte[] low = OfficeWebpCodec.Encode(source, Lossy(15));
        byte[] high = OfficeWebpCodec.Encode(source, Lossy(100));
        Assert.True(OfficeWebpCodec.TryDecode(low, out OfficeRasterImage? lowImage));
        Assert.True(OfficeWebpCodec.TryDecode(high, out OfficeRasterImage? highImage));
        Assert.True(ColorError(source, highImage!) < ColorError(source, lowImage!));
        Assert.True(high.Length > low.Length);
        Assert.True(HasChunk(high, "VP8 "));
        Assert.False(HasChunk(high, "VP8L"));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void FilteredAlphaUsesTheWebpFirstRowAndFirstColumnBoundaryRules(bool horizontal) {
        var image = new OfficeRasterImage(16, 16);
        for (int y = 0; y < 16; y++) {
            for (int x = 0; x < 16; x++) {
                byte alpha = ((horizontal ? x : y) & 1) == 0 ? (byte)24 : (byte)216;
                image.SetPixel(x, y, OfficeColor.FromRgba(40, 120, 180, alpha));
            }
        }
        byte[] bytes = OfficeWebpCodec.Encode(image, Lossy(85));
        Assert.True(OfficeWebpCodec.TryDecode(bytes, out OfficeRasterImage? result));
        for (int y = 0; y < 16; y++) {
            for (int x = 0; x < 16; x++) Assert.Equal(image.GetPixel(x, y).A, result!.GetPixel(x, y).A);
        }
    }

    [Fact]
    public void GenericWebpEncodingCarriesModeDensityAndCallerOwnedStream() {
        OfficeRasterImage source = CreateImage(33, 25, withAlpha: true);
        var options = new OfficeRasterEncodingOptions {
            Webp = Lossy(85), DpiX = 144, DpiY = 120, WriteResolutionMetadata = true
        };
        options.Webp.WritePhysicalResolution = true;
        byte[] bytes = OfficeRasterImageEncoder.Encode(source, OfficeImageExportFormat.Webp, options);
        using var stream = new MemoryStream();
        OfficeRasterImageEncoder.EncodeTo(source, OfficeImageExportFormat.Webp, stream, options);
        Assert.Equal(bytes, stream.ToArray());
        Assert.True(stream.CanWrite);
        OfficeImageInfo metadata = OfficeImageReader.Identify(bytes);
        Assert.Equal(144D, metadata.DpiX, 6);
        Assert.Equal(120D, metadata.DpiY, 6);
        Assert.Equal(0x18, bytes[20]); // VP8X alpha and Exif feature flags.
    }

    [Fact]
    public void CancellationDuringCompressionDoesNotWriteAContainer() {
        using var cancellation = new CancellationTokenSource();
        using var stream = new MemoryStream();
        int checkpoints = 0;
        Assert.Throws<OperationCanceledException>(() => OfficeWebpCodec.EncodeTo(
            CreateImage(129, 97, withAlpha: true), stream, Lossy(85), null, null, cancellation.Token,
            _ => { if (++checkpoints == 10) cancellation.Cancel(); }));
        Assert.Equal(0L, stream.Length);
        Assert.True(stream.CanWrite);
    }

    [Fact]
    public void OutputBudgetStopsWritesBeforeCrossingTheCallersLimit() {
        using var stream = new MemoryStream();
        Assert.Throws<OfficeImageExportBatchLimitException>(() => OfficeRasterImageEncoder.EncodeTo(
            CreateImage(33, 25, withAlpha: true), OfficeImageExportFormat.Webp, stream,
            new OfficeRasterEncodingOptions { Webp = Lossy(85) }, 32));
        Assert.InRange(stream.Length, 0L, 32L);
    }

    [Fact]
    public void UnsupportedVp8WidthIsRejectedInsteadOfChangingCompressionMode() {
        var image = new OfficeRasterImage(16384, 1, OfficeColor.White);
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeWebpCodec.Encode(image, Lossy(85)));
        Assert.True(OfficeWebpCodec.TryDecode(OfficeWebpCodec.Encode(image), out OfficeRasterImage? lossless));
        Assert.Equal(16384, lossless!.Width);
    }

    [Theory]
    [InlineData(OfficeWebpEncodingMode.Lossy)]
    [InlineData(OfficeWebpEncodingMode.Lossless)]
    public void WorkingSetGuardIncludesRetainedCallerBuffersBeforeWriting(OfficeWebpEncodingMode mode) {
        using var stream = new MemoryStream();
        OfficeWebpEncodeOptions options = Lossy(85);
        options.Mode = mode;
        options.RetainedManagedBytes = OfficeRasterGuards.MaximumDecodedBytes;
        Assert.Throws<ArgumentException>(() => OfficeWebpCodec.EncodeTo(CreateImage(1, 1, false), stream, options));
        Assert.Equal(0L, stream.Length);
    }

    [Theory]
    [InlineData(OfficeWebpEncodingMode.Lossy, false)]
    [InlineData(OfficeWebpEncodingMode.Lossy, true)]
    [InlineData(OfficeWebpEncodingMode.Lossless, false)]
    [InlineData(OfficeWebpEncodingMode.Lossless, true)]
    public void AppendingRejectsTransientBackingGrowthBeforeWriting(OfficeWebpEncodingMode mode, bool generic) {
        const int capacity = 1024 * 1024;
        using var stream = new ObservedMemoryStream(capacity);
        stream.SetLength(capacity);
        stream.Position = capacity;
        var options = new OfficeWebpEncodeOptions {
            Mode = mode,
            // The existing caller buffers leave room for the original backing array,
            // but not for it and the doubled backing array during the first append.
            RetainedManagedBytes = OfficeRasterGuards.MaximumDecodedBytes - capacity * 2L
        };
        var image = new OfficeRasterImage(1, 1, OfficeColor.White);
        Assert.Throws<ArgumentException>(() => {
            if (generic) {
                OfficeRasterImageEncoder.EncodeTo(image, OfficeImageExportFormat.Webp, stream,
                    new OfficeRasterEncodingOptions { Webp = options }, 4096);
            } else {
                OfficeWebpCodec.EncodeTo(image, stream, options);
            }
        });
        Assert.Equal(0, stream.Writes);
        Assert.Equal(capacity, stream.Capacity);
        Assert.Equal(capacity, stream.Length);
        Assert.Equal(capacity, stream.Position);
        Assert.True(stream.CanWrite);
    }

    [Theory]
    [InlineData(OfficeWebpEncodingMode.Lossy, false)]
    [InlineData(OfficeWebpEncodingMode.Lossy, true)]
    [InlineData(OfficeWebpEncodingMode.Lossless, false)]
    [InlineData(OfficeWebpEncodingMode.Lossless, true)]
    public void PositionedExpandableAndFixedBuffersPreservePrefixAndEncodedImage(
        OfficeWebpEncodingMode mode, bool fixedBuffer) {
        var image = new OfficeRasterImage(1, 1, OfficeColor.White);
        var options = new OfficeWebpEncodeOptions { Mode = mode };
        using var expectedStream = new MemoryStream();
        OfficeWebpCodec.EncodeTo(image, expectedStream, options);
        byte[] encoded = expectedStream.ToArray();
        byte[] prefix = { 11, 22, 33, 44 };
        foreach (bool generic in new[] { false, true }) {
            using MemoryStream stream = fixedBuffer
                ? new MemoryStream(new byte[encoded.Length + prefix.Length + 8], 0,
                    encoded.Length + prefix.Length + 8, writable: true, publiclyVisible: true)
                : new MemoryStream(2);
            stream.SetLength(0);
            stream.Write(prefix, 0, prefix.Length);
            if (generic) {
                OfficeRasterImageEncoder.EncodeTo(image, OfficeImageExportFormat.Webp, stream,
                    new OfficeRasterEncodingOptions { Webp = options }, encoded.Length);
            } else {
                OfficeWebpCodec.EncodeTo(image, stream, options);
            }
            byte[] combined = stream.ToArray();
            Assert.Equal(encoded.Length + prefix.Length, combined.Length);
            for (int index = 0; index < prefix.Length; index++) Assert.Equal(prefix[index], combined[index]);
            var payload = new byte[encoded.Length];
            Buffer.BlockCopy(combined, prefix.Length, payload, 0, payload.Length);
            Assert.Equal(encoded, payload);
            Assert.True(OfficeWebpCodec.TryDecode(payload, out OfficeRasterImage? decoded));
            Assert.Equal(OfficeColor.White, decoded!.GetPixel(0, 0));
            Assert.True(stream.CanWrite);
        }
    }

    [Theory]
    [InlineData(OfficeWebpEncodingMode.Lossy)]
    [InlineData(OfficeWebpEncodingMode.Lossless)]
    public void CallerOwnedBackingDoesNotReserveAnUnrequestedFinalCopy(OfficeWebpEncodingMode mode) {
        const int capacity = 1024 * 1024;
        var image = new OfficeRasterImage(1, 1, OfficeColor.White);
        using var expectedStream = new MemoryStream();
        OfficeWebpCodec.EncodeTo(image, expectedStream, new OfficeWebpEncodeOptions { Mode = mode });
        byte[] encoded = expectedStream.ToArray();
        foreach (bool generic in new[] { false, true }) {
            using var stream = new MemoryStream(capacity);
            stream.SetLength(capacity);
            var options = new OfficeWebpEncodeOptions {
                Mode = mode,
                RetainedManagedBytes = OfficeRasterGuards.MaximumDecodedBytes - capacity * 2L
            };
            if (generic) {
                OfficeRasterImageEncoder.EncodeTo(image, OfficeImageExportFormat.Webp, stream,
                    new OfficeRasterEncodingOptions { Webp = options }, encoded.Length);
            } else {
                OfficeWebpCodec.EncodeTo(image, stream, options);
            }
            var payload = new byte[encoded.Length];
            Buffer.BlockCopy(stream.GetBuffer(), 0, payload, 0, payload.Length);
            Assert.Equal(encoded, payload);
            Assert.Equal(encoded.Length, stream.Position);
            Assert.Equal(capacity, stream.Capacity);
            Assert.Equal(capacity, stream.Length);
        }
    }

    private sealed class ObservedMemoryStream : MemoryStream {
        internal ObservedMemoryStream(int capacity) : base(capacity) { }
        internal int Writes { get; private set; }
        public override void Write(byte[] buffer, int offset, int count) {
            Writes++;
            base.Write(buffer, offset, count);
        }
        public override void WriteByte(byte value) {
            Writes++;
            base.WriteByte(value);
        }
    }

    [Theory]
    [InlineData(0)]
    [InlineData(101)]
    public void LossyQualityOutsideItsDocumentedRangeIsRejected(int quality) {
        Assert.Throws<ArgumentOutOfRangeException>(() => OfficeWebpCodec.Encode(CreateImage(1, 1, false), Lossy(quality)));
    }

    private static OfficeWebpEncodeOptions Lossy(int quality) => new OfficeWebpEncodeOptions {
        Mode = OfficeWebpEncodingMode.Lossy, Quality = quality
    };

    private static OfficeRasterImage CreateImage(int width, int height, bool withAlpha) {
        var image = new OfficeRasterImage(width, height);
        for (int y = 0; y < height; y++) {
            for (int x = 0; x < width; x++) {
                byte red = (byte)Math.Max(0, Math.Min(255, 110 + 70 * Math.Sin(x / 37.0) + 60 * Math.Cos(y / 31.0)));
                byte green = (byte)Math.Max(0, Math.Min(255, 100 + 75 * Math.Cos(x / 19.0 + y / 47.0) + 65 * Math.Sin(y / 29.0)));
                byte blue = (byte)Math.Max(0, Math.Min(255, 120 + 60 * Math.Sin(x / 29.0 + y / 13.0) + 60 * Math.Cos(y / 23.0)));
                image.SetPixel(x, y, OfficeColor.FromRgba(red, green, blue, withAlpha ? (byte)(x * 7 + y * 13) : (byte)255));
            }
        }
        return image;
    }

    private static double ColorError(OfficeRasterImage source, OfficeRasterImage result) {
        byte[] expected = source.GetPixels(), actual = result.GetPixels();
        double squaredError = 0;
        for (int i = 0; i < expected.Length; i++) {
            if ((i & 3) != 3) squaredError += (expected[i] - actual[i]) * (expected[i] - actual[i]);
        }
        return squaredError / (source.Width * source.Height * 3D);
    }

    private static bool HasChunk(byte[] bytes, string name) {
        for (int offset = 12; offset + 8 <= bytes.Length;) {
            if (System.Text.Encoding.ASCII.GetString(bytes, offset, 4) == name) return true;
            int length = bytes[offset + 4] | bytes[offset + 5] << 8 | bytes[offset + 6] << 16 | bytes[offset + 7] << 24;
            offset += 8 + length + (length & 1);
        }
        return false;
    }
}
