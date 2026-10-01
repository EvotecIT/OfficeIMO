using System;
using System.IO;
using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DrawingWebpVp8Tests {
    [Fact]
    public void ReservedColorSpaceIsRejectedButValidNoClampFlagIsAccepted() {
        var reservedColor = new OfficeVp8BoolDecoder(new OfficeByteView(new byte[] { 128, 0 }));
        Assert.False(OfficeVp8Decoder.TryReadControlHeader(reservedColor, out _));

        var noClampRequired = new OfficeVp8BoolDecoder(new OfficeByteView(new byte[] { 64, 0 }));
        Assert.True(OfficeVp8Decoder.TryReadControlHeader(noClampRequired, out var header));
        Assert.Equal(0, header.ColorSpace);
        Assert.Equal(1, header.ClampType);
    }

    [Fact]
    public void ArithmeticReadFailsOnTheSymbolThatExhaustsPadding() {
        var decoder = new OfficeVp8BoolDecoder(new OfficeByteView(new byte[] { 0, 0 }));
        for (int index = 0; index < 16; index++) {
            Assert.True(decoder.TryReadBool(128, out _));
        }

        Assert.False(decoder.TryReadBool(128, out _));
    }

    [Fact]
    public void ZeroBaseFilterLevelStillAppliesPositiveReferenceDelta() {
        var loopFilter = new OfficeVp8LoopFilter(
            filterType: 1, level: 0, sharpness: 0, deltaEnabled: true, deltaUpdate: true,
            refDeltas: new[] { 10, 0, 0, 0 }, refDeltasUpdated: new[] { true, false, false, false },
            modeDeltas: new int[4], modeDeltasUpdated: new bool[4]);
        var segmentation = new OfficeVp8Segmentation(
            enabled: false, updateMap: false, updateData: false, absoluteDeltas: false,
            quantizerDeltas: Array.Empty<int>(), filterDeltas: Array.Empty<int>(),
            segmentProbabilities: Array.Empty<int>());
        var macroblocks = new[] {
            new OfficeVp8MacroblockHeader(0, 0, 0, 0, false, 0, 0, false, Array.Empty<int>()),
            new OfficeVp8MacroblockHeader(1, 1, 0, 0, false, 0, 0, false, Array.Empty<int>())
        };
        byte[] luma = new byte[32 * 16];
        for (int y = 0; y < 16; y++) {
            for (int x = 0; x < 32; x++) luma[y * 32 + x] = x < 16 ? (byte)100 : (byte)110;
        }

        OfficeVp8Decoder.ApplyLoopFilter(loopFilter, segmentation, macroblocks,
            new[] { false, false }, 32, 16, luma, new byte[16 * 8], new byte[16 * 8],
            16, 8, isKeyframe: true, cancellationToken: CancellationToken.None);

        Assert.True(luma[15] > 100);
        Assert.True(luma[16] < 110);
    }

    [Fact]
    public void VerticalLeftPredictionUsesSpecifiedLastTwoTopTriples() {
        // RFC 6386 section 12.3 assigns B[2][3] to A[4..6] and
        // B[3][3] to A[5..7]; the last two pixels break the diagonal pattern.
        byte[] plane = new byte[16 * 16];
        for (int index = 0; index < 8; index++) {
            plane[3 * 16 + 4 + index] = (byte)((index + 1) * 10);
        }
        byte[] predicted = new byte[16];

        OfficeVp8Prediction.PredictSubblock(plane, 16, 16, 4, 4, 7, predicted, new OfficeVp8DecodeScratch());

        Assert.Equal((byte)60, predicted[2 * 4 + 3]);
        Assert.Equal((byte)70, predicted[3 * 4 + 3]);
    }

    [Theory]
    [InlineData(20)]
    [InlineData(24)]
    [InlineData(28)]
    public void RightEdgeSubblocksReuseTopRightPixelsAboveMacroblock(int y) {
        // RFC 6386 section 12.3: blocks 7, 11, and 15 reuse block 3's
        // top-right pixels from the row above the entire macroblock.
        byte[] plane = new byte[32 * 32];
        plane[15 * 32 + 16] = 50;
        plane[15 * 32 + 17] = 60;
        plane[(y - 1) * 32 + 15] = 40;
        plane[(y - 1) * 32 + 16] = 200; // Not yet reconstructed.
        byte[] predicted = new byte[16];

        OfficeVp8Prediction.PredictSubblock(plane, 32, 32, 12, y, 4, predicted, new OfficeVp8DecodeScratch());

        Assert.Equal((byte)50, predicted[3]);
    }

    // Independently encoded and decoded with Pillow/libwebp (quality 80, method 6).
    // These fixtures exercise directional prediction, coefficients and odd canvas padding.
    [Theory]
    [InlineData("independent-vp8-pattern", 48, 32)]
    [InlineData("independent-vp8-odd-color", 65, 49)]
    public void OwnedDecoderMatchesIndependentPixels(string name, int width, int height) {
        byte[] encoded = ReadFixture(name + ".webp");
        byte[] expected = ReadFixture(name + ".rgba");
        Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out var image));
        Assert.True(OfficeRasterContainerInspector.TryInspect(encoded, out var container));
        Assert.Equal((width, height), (container!.CanvasWidth, container.CanvasHeight));
        Assert.Equal(width, image!.Width);
        Assert.Equal(height, image.Height);
        byte[] actual = image.GetPixels();
        Assert.Equal(expected.Length, actual.Length);
        long totalError = 0;
        for (int index = 0; index < actual.Length; index++) {
            int error = Math.Abs(actual[index] - expected[index]);
            Assert.InRange(error, 0, 3);
            totalError += error;
        }
        Assert.True(totalError <= actual.Length, "Mean channel error exceeds one.");
        Assert.True(OfficeImageReader.TryValidateContent(encoded, null, out _));
    }

    [Theory]
    [InlineData("UklGRjwAAABXRUJQVlA4IDAAAADQAQCdASoQABAAAUAmJaACdLoB+AADsAD+8ut//NgVzXPv9//S4P0uD9Lg/9KQAAA=", 255, 1, 0)]
    [InlineData("UklGRiQAAABXRUJQVlA4IBgAAABQAQCdASoQABAAAUAmJaQABHQAAP4AAAA=", 127, 127, 127)]
    public void OwnedDecoderPreservesIndependentSolidColors(string encoded, int red, int green, int blue) {
        Assert.True(OfficeWebpCodec.TryDecode(Convert.FromBase64String(encoded), out var image));
        Assert.Equal(16, image!.Width);
        Assert.Equal(16, image.Height);
        byte[] pixels = image.GetPixels();
        for (int offset = 0; offset < pixels.Length; offset += 4) {
            Assert.InRange((int)pixels[offset], Math.Max(0, red - 1), Math.Min(255, red + 1));
            Assert.InRange((int)pixels[offset + 1], Math.Max(0, green - 1), Math.Min(255, green + 1));
            Assert.InRange((int)pixels[offset + 2], Math.Max(0, blue - 1), Math.Min(255, blue + 1));
            Assert.Equal(255, pixels[offset + 3]);
        }
    }

    [Fact]
    public void PixelLimitAndCancellationStopOwnedDecode() {
        byte[] encoded = ReadFixture("independent-vp8-pattern.webp");
        Assert.False(OfficeRasterImageDecoder.TryDecode(encoded,
            new OfficeRasterDecodeOptions { MaximumDecodedPixels = 100 }, out _, out _));
        using var cancellation = new CancellationTokenSource();
        cancellation.Cancel();
        Assert.Throws<OperationCanceledException>(() => OfficeRasterImageDecoder.TryDecode(encoded,
            new OfficeRasterDecodeOptions { CancellationToken = cancellation.Token }, out _, out _));
    }

    [Fact]
    public void AggregateDecodeBudgetRejectsLargeCanvasBeforeAllocatingPlanes() {
        byte[] encoded = ReadFixture("independent-vp8-pattern.webp");
        // 49 million pixels fit the pixel ceiling but planes plus RGBA exceed the shared byte budget.
        encoded[26] = encoded[28] = (byte)(7000 & 255);
        encoded[27] = encoded[29] = (byte)(7000 >> 8);
        Assert.False(OfficeWebpCodec.TryDecode(encoded, out _));
    }

#if NET8_0_OR_GREATER
    [Fact]
#if DRAWING_PERFORMANCE_EVIDENCE
    [Trait("Category", "Performance")]
    [Trait("Category", "ResourcePerformanceEvidence")]
#endif
    public void LargeVp8DecodeKeepsAllocationsNearItsRetainedImageBudget() {
        // Independently produced with Pillow/libwebp from a 512x512 gradient pattern.
        // The output itself needs 1 MiB; per-block coefficient/transform arrays
        // would add several more MiB of short-lived allocations.
        byte[] encoded = ReadFixture("independent-vp8-allocation.webp");
#if DRAWING_PERFORMANCE_EVIDENCE
        Assert.True(OfficeWebpCodec.TryDecode(encoded, out _)); // JIT and codec warmup.

        long before = GC.GetAllocatedBytesForCurrentThread();
#endif
        bool decoded = OfficeWebpCodec.TryDecode(encoded, out OfficeRasterImage? image);
#if DRAWING_PERFORMANCE_EVIDENCE
        long allocated = GC.GetAllocatedBytesForCurrentThread() - before;
#endif

        Assert.True(decoded);
        Assert.Equal(512, image!.Width);
        Assert.Equal(512, image.Height);
#if DRAWING_PERFORMANCE_EVIDENCE
        Assert.True(allocated <= 4_000_000, $"VP8 decode allocated {allocated:N0} bytes.");
#endif
    }
#endif

    [Fact]
    public void TruncatedControlPartitionIsRejected() {
        byte[] encoded = ReadFixture("independent-vp8-pattern.webp");
        Array.Resize(ref encoded, 24);
        encoded[4] = 16;
        encoded[5] = encoded[6] = encoded[7] = 0;
        Assert.False(OfficeWebpCodec.TryDecode(encoded, out _));
    }

    [Theory]
    [InlineData(true)]
    [InlineData(false)]
    public void OversizedContainerLengthReturnsFalseFromPublicDecoder(bool riffLength) {
        byte[] encoded = ReadFixture("independent-vp8-pattern.webp");
        int offset = riffLength ? 4 : 16;
        for (int index = 0; index < 4; index++) encoded[offset + index] = 0xff;

        Assert.False(OfficeWebpCodec.TryDecode(encoded, out _));
    }

    [Theory]
    [InlineData(false)]
    [InlineData(true)]
    public void ExhaustedArithmeticPartitionCannotProduceAnImage(bool tokenPartition) {
        byte[] encoded;
        if (tokenPartition) {
            encoded = ReadFixture("independent-vp8-pattern.webp");
            int controlLength = (encoded[20] | encoded[21] << 8 | encoded[22] << 16) >> 5;
            int payloadLength = 10 + controlLength + 2;
            Array.Resize(ref encoded, 20 + payloadLength + (payloadLength & 1));
            Array.Clear(encoded, 30 + controlLength, encoded.Length - 30 - controlLength);
            WriteLength(encoded, 16, payloadLength);
            WriteLength(encoded, 4, encoded.Length - 8);
        } else {
            // Valid container/header, but only two zero bytes in each partition.
            encoded = new byte[] {82,73,70,70,26,0,0,0,87,69,66,80,86,80,56,32,
                14,0,0,0,80,0,0,157,1,42,16,0,16,0,0,0,0,0};
        }
        Assert.True(OfficeImageReader.TryIdentifyByContent(encoded, null, out _));
        Assert.False(OfficeRasterContainerInspector.TryInspect(encoded, out _));
        Assert.False(OfficeWebpCodec.TryDecode(encoded, out _));
        Assert.False(OfficeRasterImageDecoder.TryDecode(encoded, out _));
        Assert.False(OfficeImageReader.TryValidateContent(encoded, null, out _));
        var drawing = new OfficeDrawing(1, 1).AddImage(encoded, "image/webp",
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, 1, 1)));
        OfficeRasterImage rendered = OfficeDrawingRasterRenderer.Render(drawing,
            new OfficeDrawingRasterRenderOptions { ImageCodec = new UnexpectedWebpCodec() });
        Assert.Equal(OfficeColor.Transparent, rendered.GetPixel(0, 0));
    }

    private sealed class UnexpectedWebpCodec : IOfficeRasterImageCodec {
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            throw new InvalidOperationException("Malformed opaque VP8 must not reach the caller codec.");
        }
    }

    [Fact]
    public void StaticLossyWebpWithSeparateAlphaPlaneUsesManagedDecoder() {
        byte[] opaque = ReadFixture("independent-vp8-pattern.webp");
        const int width = 48;
        const int height = 32;
        int alphaLength = 1 + width * height;
        byte[] encoded = new byte[12 + 18 + 8 + alphaLength + (alphaLength & 1) + opaque.Length - 12];
        Array.Copy(opaque, encoded, 12);
        WriteLength(encoded, 4, encoded.Length - 8);
        Array.Copy(System.Text.Encoding.ASCII.GetBytes("VP8X"), 0, encoded, 12, 4);
        WriteLength(encoded, 16, 10);
        encoded[20] = 0x10; // Alpha flag.
        encoded[24] = width - 1;
        encoded[27] = height - 1;
        int alphaChunk = 30;
        Array.Copy(System.Text.Encoding.ASCII.GetBytes("ALPH"), 0, encoded, alphaChunk, 4);
        WriteLength(encoded, alphaChunk + 4, alphaLength);
        for (int index = alphaChunk + 9; index < alphaChunk + 8 + alphaLength; index++) encoded[index] = 255;
        Array.Copy(opaque, 12, encoded, alphaChunk + 8 + alphaLength + (alphaLength & 1), opaque.Length - 12);

        Assert.True(OfficeImageReader.TryIdentifyByContent(encoded, null, out _));
        Assert.True(OfficeRasterImageDecoder.TryDecode(encoded, out var expected));
        var codec = new SolidWebpCodec();
        var drawing = new OfficeDrawing(width, height).AddImage(encoded, "image/webp",
            new OfficeImageProjection(new OfficeImagePlacement(0, 0, width, height)));
        OfficeRasterImage rendered = OfficeDrawingRasterRenderer.Render(drawing,
            new OfficeDrawingRasterRenderOptions { ImageCodec = codec });
        Assert.Equal(0, codec.Calls);
        Assert.Equal(expected!.GetPixel(13, 11), rendered.GetPixel(13, 11));
    }

    private sealed class SolidWebpCodec : IOfficeRasterImageCodec {
        internal int Calls;
        public bool TryDecode(byte[] encodedBytes, string? contentType, out OfficeRasterImage? image) {
            Calls++;
            image = new OfficeRasterImage(48, 32, OfficeColor.Red);
            return true;
        }
    }

    [Fact]
    public void InverseVp8TransformsUseReferenceSigned16IntermediateStorage() {
        // RFC 6386 section 14 stores dequantized coefficients and each transform
        // stage in signed 16-bit arrays. These extreme reference vectors diverge
        // if a stage is retained as an unbounded 32-bit integer.
        var coefficients = new int[16];
        for (int index = 0; index < coefficients.Length; index++) coefficients[index] = 32767;
        Assert.Equal(new[] {
            -2402, 478, -478, -95, -12062, 2399, -2399, -477,
            12062, -2399, 2399, 477, 2400, -477, 477, 95
        }, OfficeVp8Transform.InverseTransform4x4(coefficients));
        Assert.Equal(new[] {
            -2, 0, 0, 0, 0, 0, 0, 0,
            0, 0, 0, 0, 0, 0, 0, 0
        }, OfficeVp8Transform.InverseWalshTransform4x4(coefficients));
    }

    private static void WriteLength(byte[] bytes, int offset, int value) {
        for (int index = 0; index < 4; index++) bytes[offset + index] = (byte)(value >> (8 * index));
    }

    private static byte[] ReadFixture(string name) {
        using Stream input = typeof(DrawingWebpVp8Tests).Assembly.GetManifestResourceStream(
            "OfficeIMO.Drawing.Tests.TestAssets.WebpVp8." + name)!;
        Assert.NotNull(input);
        using var output = new MemoryStream();
        input.CopyTo(output);
        return output.ToArray();
    }
}
