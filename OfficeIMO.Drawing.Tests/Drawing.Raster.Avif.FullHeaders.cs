using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Native full still headers protect decoding of non-reduced Main color/grayscale items and alpha.</summary>
public sealed class DrawingAvifFullHeaderTests {
    public static IEnumerable<object[]> Cases() {
        foreach (int depth in new[] { 8, 10 })
            foreach (string format in new[] { "420", "mono" })
                foreach (string range in new[] { "full", "limited" })
                    foreach (string alpha in new[] { "", "-alpha" })
                        yield return new object[] { $"avif-main{depth}-fullheader-{format}-{range}{alpha}" };
    }

    [Theory]
    [MemberData(nameof(Cases))]
    public void IndependentFullHeaderStillItemsDecodeReferencePixelsAndRestoreStreams(string name) {
        string path = Path.Combine(AppContext.BaseDirectory, "TestAssets", "Avif", name + ".avif");
        byte[] bytes = File.ReadAllBytes(path), saved = (byte[])bytes.Clone();
        byte[] expected = File.ReadAllBytes(Path.ChangeExtension(path, ".rgba"));
        Assert.True(OfficeRasterImageDecoder.TryDecode(bytes, new OfficeRasterDecodeOptions(), out var image, out var info));
        Assert.NotNull(image);
        Assert.Equal(49, image.Width); Assert.Equal(33, image.Height);
        Assert.Equal(OfficeImageFormat.Avif, info.Format); Assert.True(info.Succeeded);
        byte[] actual = image.GetPixels(); Assert.Equal(expected.Length, actual.Length);
        for (int i = 0; i < actual.Length; i++) Assert.InRange(Math.Abs(actual[i] - expected[i]), 0, i % 4 == 3 ? 0 : 1);
        Assert.Equal(saved, bytes);
        using var stream = new MemoryStream(bytes);
        Assert.True(OfficeRasterImageDecoder.TryDecode(stream, out var streamed));
        Assert.Equal(0, stream.Position); Assert.Equal(actual, streamed!.GetPixels());
        var options = new OfficeRasterDecodeOptions();
        Assert.True(OfficeAvifContainerReader.TryRead(bytes, options, out var container));
        using var planes = new BinaryReader(File.OpenRead(Path.ChangeExtension(path, ".yuv16")));
        var items = container!.Alpha == null ? new[] { container.Color } : new[] { container.Color, container.Alpha };
        foreach (var item in items) {
            Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes, item!, options, out var sequence));
            Assert.False(sequence!.ReducedStillHeader);
            Assert.True(OfficeAv1StillFrameReader.TryRead(bytes, item!, sequence, options, out var frame));
            var reconstructed = OfficeAv1FrameReconstructor.Decode(bytes, sequence, frame!, options, OfficeAv1ReconstructionStage.Restored);
            for (int p = 0; p < reconstructed.PlaneCount; p++) {
                int width = p == 0 ? 49 : 25, height = p == 0 ? 33 : 17;
                for (int y = 0; y < height; y++) for (int x = 0; x < width; x++)
                    Assert.Equal((int)planes.ReadUInt16(), reconstructed.Value(p, x, y));
            }
        }
        Assert.Equal(planes.BaseStream.Length, planes.BaseStream.Position);
    }

    [Fact]
    public void FullHeaderStillCannotShowAnExistingInterOrHiddenFrameOrBorrowFollowingBytes() {
        byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Avif", "avif-main10-fullheader-420-full-alpha.avif"));
        var options = new OfficeRasterDecodeOptions();
        Assert.True(OfficeAvifContainerReader.TryRead(bytes, options, out var container));
        foreach (var item in new[] { container!.Color, container.Alpha! }) {
            Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes, item, options, out var sequence));
            Assert.False(sequence!.ReducedStillHeader);
            var (sizeOffset, payloadOffset) = FindSequence(bytes, item);
            foreach (int mask in new[] { 4, 1, 16 }) { // Timing model, multiple operating points, non-still sequence.
                byte[] unsupported = (byte[])bytes.Clone(); unsupported[payloadOffset] ^= (byte)mask;
                Assert.False(OfficeAvifCodec.TryDecode(unsupported, options, out _, out bool unsupportedEligible));
                Assert.False(unsupportedEligible);
            }
            for (int truncated = 0; truncated < bytes[sizeOffset]; truncated++) {
                byte[] malformed = (byte[])bytes.Clone(); malformed[sizeOffset] = (byte)truncated;
                Assert.False(OfficeAv1StillSequenceReader.TryRead(malformed, item, options, out _));
            }
            Assert.True(OfficeAv1StillFrameReader.TryRead(bytes, item, sequence, options, out var frame));
            foreach (int mask in new[] { 128, 32, 16 }) { // show_existing_frame, inter-frame type, hidden frame.
                byte[] malformed = (byte[])bytes.Clone(); malformed[sequence.FrameOffset] ^= (byte)mask;
                Assert.False(OfficeAvifCodec.TryDecode(malformed, options, out _, out bool eligible));
                Assert.False(eligible);
            }
            int length = sequence.FrameLength;
            for (int truncated = 0; truncated <= frame!.HeaderBytes; truncated++) {
                sequence.FrameLength = truncated;
                Assert.False(OfficeAv1StillFrameReader.TryRead(bytes, item, sequence, options, out _));
            }
            sequence.FrameLength = length;
        }
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, new OfficeRasterDecodeOptions { MaximumEncodedBytes = bytes.Length - 1 }, out _, out _));
        Assert.False(OfficeRasterImageDecoder.TryDecode(bytes, new OfficeRasterDecodeOptions { MaximumDecodedPixels = 1616 }, out _, out _));
        Assert.False(OfficeAvifCodec.TryDecode(bytes, new OfficeRasterDecodeOptions { RetainedManagedBytes = OfficeRasterGuards.MaximumDecodedBytes - 1 }, out _, out bool retainedEligible));
        Assert.False(retainedEligible);
        Assert.False(OfficeAvifCodec.TryDecode(bytes, new OfficeRasterDecodeOptions { MaximumInspectionWorkPixels = 1 }, out _, out bool workEligible));
        Assert.False(workEligible);
        Assert.Throws<OperationCanceledException>(() => OfficeRasterImageDecoder.TryDecode(bytes,
            new OfficeRasterDecodeOptions { CancellationToken = new System.Threading.CancellationToken(true) }, out _, out _));
    }

    private static (int SizeOffset, int PayloadOffset) FindSequence(byte[] bytes, OfficeAvifImageItem item) {
        for (int p = item.Offset; p < item.Offset + item.Length;) {
            int header = bytes[p++];
            Assert.Equal(2, header & 6); // Native fixtures use size fields, without OBU extensions.
            int sizeOffset = p, size = bytes[p++];
            Assert.InRange(size, 0, 127); // This fixture's sequence/delimiter sizes fit one LEB byte.
            if ((header >> 3 & 15) == 1) return (sizeOffset, p);
            p += size;
        }
        throw new InvalidDataException("Native fixture has no sequence OBU");
    }

    [Theory]
    [InlineData(0, 2, true)]
    [InlineData(1, 0, false)]
    [InlineData(2, 2, false)]
    public void FullFrameFieldChoicesKeepSizeAndTileBoundaryAligned(int screenTools, int integerMv, bool disableCdf) {
        // A lossless monochrome syntax fixture covers fields absent from the native producer's defaults.
        var bits = new HeaderWriter();
        bits.Write(1, 4); // No show-existing; KEY_FRAME; show_frame.
        bits.Write(disableCdf ? 1 : 0, 1);
        if (screenTools == 2) bits.Write(1, 1);
        if (screenTools != 0 && integerMv == 2) bits.Write(0, 1);
        bits.Write(65535, 16); // Current frame ID, maximum supported field length.
        bits.Write(1, 1); bits.Write(255, 8); // Size override, order hint.
        bits.Write(48, 6); bits.Write(32, 6); // Actual 49x33 in a 64x64 sequence.
        bits.Write(0, 1); // Same render size.
        if (screenTools != 0) bits.Write(0, 1); // No intra-block copy.
        if (!disableCdf) bits.Write(1, 1); // End-of-frame CDF update disabled.
        bits.Write(1, 1); bits.Write(0, 8); bits.Write(0, 3); // Uniform tile, Q=0, no DC delta/matrix/segmentation.
        bits.Write(0, 1); bits.Align(); // Reduced transform set false; no filter/tx-mode fields for lossless.
        byte[] payload = bits.Bytes.Concat(new byte[] { 123 }).ToArray();
        var item = new OfficeAvifImageItem(1, 49, 33, 0, payload.Length, new byte[] { 129, 0, 28, 0 }, false, null);
        var sequence = new OfficeAv1StillSequence {
            ReducedStillHeader = false, ForceScreenContentTools = screenTools, ForceIntegerMv = integerMv,
            FrameIdBits = 16, OrderHintBits = 8, WidthBits = 6, HeightBits = 6,
            MaximumWidth = 64, MaximumHeight = 64, Monochrome = true, FrameLength = payload.Length
        };
        Assert.True(OfficeAv1StillFrameReader.TryRead(payload, item, sequence, new OfficeRasterDecodeOptions(), out var frame));
        Assert.True(frame!.CodedLossless); Assert.Equal(disableCdf, frame.DisableCdfUpdate);
        Assert.Equal(screenTools != 0, frame.AllowScreenContentTools);
        Assert.Equal(49, frame.UpscaledWidth); Assert.Equal(33, frame.Height);
        Assert.Equal(123, payload[Assert.Single(frame.Tiles).Offset]);
    }

    private sealed class HeaderWriter {
        internal List<byte> Bytes { get; } = new List<byte>();
        private int _position;
        internal void Write(int value, int count) {
            for (int i = count - 1; i >= 0; i--) {
                if ((_position & 7) == 0) Bytes.Add(0);
                Bytes[Bytes.Count - 1] |= (byte)(((value >> i) & 1) << (7 - (_position & 7)));
                _position++;
            }
        }
        internal void Align() { while ((_position & 7) != 0) Write(0, 1); }
    }
}
