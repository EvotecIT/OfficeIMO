using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Independent libavif/FFmpeg headers plus bounded syntax fixtures protect entry into AV1 entropy payloads.</summary>
public sealed class DrawingAv1FrameTests {
    [Theory]
    [InlineData("avif-opaque", false, 12, 31, 32)]
    [InlineData("avif-alpha", true, 5, 43, 24)]
    public void FrozenColorAndAlphaHeadersMatchIndependentTrace(string name, bool alpha, int headerBytes, int tileBytes, int q) {
        var (bytes, item, sequence) = ReadFixture(name, alpha);
        Assert.True(OfficeAv1StillFrameReader.TryRead(bytes, item, sequence, new OfficeRasterDecodeOptions(), out var frame));
        Assert.NotNull(frame);
        Assert.Equal(49, frame.Width);
        Assert.Equal(33, frame.Height);
        Assert.Equal(49, frame.RenderWidth);
        Assert.Equal(33, frame.RenderHeight);
        Assert.Equal(14, frame.MiCols);
        Assert.Equal(10, frame.MiRows);
        Assert.Equal(headerBytes, frame.HeaderBytes);
        Assert.Equal(0, frame.TileGroupHeaderBytes);
        Assert.Equal(q, frame.BaseQIndex);
        Assert.False(frame.CodedLossless);
        Assert.False(frame.DisableCdfUpdate);
        Assert.Equal(!alpha, frame.AllowScreenContentTools);
        Assert.False(frame.AllowIntraBlockCopy);
        Assert.Equal(0, frame.DeltaQYDc);
        Assert.Equal(alpha ? 0 : -2, frame.DeltaQUDc);
        Assert.Equal(frame.DeltaQUDc, frame.DeltaQUAc);
        Assert.Equal(frame.DeltaQUDc, frame.DeltaQVDc);
        Assert.Equal(frame.DeltaQUDc, frame.DeltaQVAc);
        Assert.Equal(!alpha, frame.UsingQMatrix);
        if (!alpha) Assert.Equal(new[] { 10, 10, 10 }, frame.QMatrixLevels);
        Assert.Equal(!alpha, frame.DeltaQPresent);
        Assert.False(frame.DeltaLoopFilterPresent);
        Assert.Equal(alpha ? new[] { 0, 0, 0, 0 } : new[] { 1, 1, 1, 1 }, frame.LoopFilterLevels);
        Assert.Equal(alpha ? 0 : 7, frame.LoopFilterSharpness);
        Assert.Equal(alpha ? 3 : 4, frame.CdefDamping);
        Assert.Equal(1, frame.TransformMode);
        var tile = Assert.Single(frame.Tiles);
        Assert.Equal(sequence.FrameOffset + headerBytes, tile.Offset);
        Assert.Equal(tileBytes, tile.Length);
        Assert.Equal(sequence.FrameOffset + sequence.FrameLength, tile.Offset + tile.Length);
        Assert.Equal(0, tile.MiRowStart);
        Assert.Equal(10, tile.MiRowEnd);
        Assert.Equal(0, tile.MiColStart);
        Assert.Equal(14, tile.MiColEnd);
    }

    [Fact]
    public void IndependentMultiTileFrameRetainsLengthsAndGeometry() {
        var (bytes, item, sequence) = ReadFixture("multitile", false);
        Assert.True(OfficeAv1StillFrameReader.TryRead(bytes, item, sequence, new OfficeRasterDecodeOptions(), out var frame));
        Assert.Equal(7, frame!.HeaderBytes);
        Assert.Equal(1, frame.TileGroupHeaderBytes);
        Assert.Equal(new[] { 0, 32, 64 }, frame.MiColStarts);
        Assert.Equal(new[] { 0, 16, 32 }, frame.MiRowStarts);
        Assert.Equal(1, frame.ContextUpdateTileId);
        Assert.Equal(2, frame.TileSizeBytes);
        Assert.Equal(112, frame.BaseQIndex);
        Assert.Equal(2, frame.TransformMode);
        Assert.Equal(new[] { 8, 8, 8, 8 }, frame.LoopFilterLevels);
        Assert.Equal(new[] { 23, 1214, 2431, 3623 }, frame.Tiles.Select(t => t.Offset - item.Offset));
        Assert.Equal(new[] { 1189, 1215, 1192, 1208 }, frame.Tiles.Select(t => t.Length));
        Assert.Equal(new[] { 0, 0, 16, 16 }, frame.Tiles.Select(t => t.MiRowStart));
        Assert.Equal(new[] { 0, 32, 0, 32 }, frame.Tiles.Select(t => t.MiColStart));
    }

    [Fact]
    public void TileSizesCannotEscapeFrameOrLeaveEmptyLastTile() {
        var (original, item, sequence) = ReadFixture("multitile", false);
        Assert.True(OfficeAv1StillFrameReader.TryRead(original, item, sequence, new OfficeRasterDecodeOptions(), out var frame));
        byte[] bytes = (byte[])original.Clone();
        int firstSize = frame!.Tiles[0].Offset - 2;
        bytes[firstSize] = bytes[firstSize + 1] = 255;
        Assert.False(OfficeAv1StillFrameReader.TryRead(bytes, item, sequence, new OfficeRasterDecodeOptions(), out _));
        bytes = (byte[])original.Clone();
        int thirdSize = frame.Tiles[2].Offset - 2;
        int consumeRemaining = sequence.FrameOffset + sequence.FrameLength - frame.Tiles[2].Offset - 1;
        bytes[thirdSize] = (byte)consumeRemaining;
        bytes[thirdSize + 1] = (byte)(consumeRemaining >> 8);
        Assert.False(OfficeAv1StillFrameReader.TryRead(bytes, item, sequence, new OfficeRasterDecodeOptions(), out _));
    }

    [Theory]
    [InlineData(128)] // Explicit partial range 0..0.
    [InlineData(152)] // Explicit full range 0..3 is also forbidden in OBU_FRAME.
    public void CombinedFrameCannotDeclareTileRanges(int groupByte) {
        var (bytes, item, sequence) = ReadFixture("multitile", false);
        Assert.True(OfficeAv1StillFrameReader.TryRead(bytes, item, sequence, new OfficeRasterDecodeOptions(), out var frame));
        bytes[sequence.FrameOffset + frame!.HeaderBytes] = (byte)groupByte;
        Assert.False(OfficeAv1StillFrameReader.TryRead(bytes, item, sequence, new OfficeRasterDecodeOptions(), out _));
    }

    [Theory]
    [InlineData("avif-opaque", false, 12, 92)]
    [InlineData("avif-alpha", true, 5, 35)]
    public void HeaderCannotBorrowFollowingBytesAndAlignmentMustBeZero(string name, bool alpha, int headerBytes, int alignmentBit) {
        var (bytes, item, sequence) = ReadFixture(name, alpha);
        int originalLength = sequence.FrameLength;
        for (int length = 0; length <= headerBytes; length++) {
            sequence.FrameLength = length;
            Assert.False(OfficeAv1StillFrameReader.TryRead(bytes, item, sequence, new OfficeRasterDecodeOptions(), out _));
        }
        sequence.FrameLength = originalLength;
        bytes[sequence.FrameOffset + alignmentBit / 8] |= (byte)(1 << (7 - alignmentBit % 8));
        Assert.False(OfficeAv1StillFrameReader.TryRead(bytes, item, sequence, new OfficeRasterDecodeOptions(), out _));
    }

    [Fact]
    public void FrameChecksActualIspeDimensionsPixelBudgetAndCancellation() {
        var (bytes, item, sequence) = ReadFixture("avif-opaque", false);
        var mismatch = new OfficeAvifImageItem(item.Id, 48, item.Height, item.Offset, item.Length,
            item.Configuration, item.Monochrome, item.ColorDescription);
        Assert.False(OfficeAv1StillFrameReader.TryRead(bytes, mismatch, sequence, new OfficeRasterDecodeOptions(), out _));
        Assert.False(OfficeAv1StillFrameReader.TryRead(bytes, item, sequence,
            new OfficeRasterDecodeOptions { MaximumDecodedPixels = 1616 }, out _));
        Assert.Throws<OperationCanceledException>(() => OfficeAv1StillFrameReader.TryRead(bytes, item, sequence,
            new OfficeRasterDecodeOptions { CancellationToken = new CancellationToken(true) }, out _));
    }

    [Fact]
    public void NonUniformLosslessTilesHaveDistinctColumnExtents() {
        // AV1 syntax fixture: 3x2 superblocks, columns 1+2, rows 1+1, full four-tile group.
        var bits = new HeaderWriter();
        bits.Write(0, 3); // CDF update, screen content, same render size.
        bits.Write(0, 1); // Nonuniform tiles.
        bits.Write(0, 1); bits.Write(1, 1); // ns(3)=0 then ns(2)=1: widths 1 and 2.
        bits.Write(0, 1); // ns(2)=0 then ns(1)=0: heights 1 and 1.
        bits.Write(2, 2); bits.Write(0, 2); // Context tile 2, one-byte lengths.
        bits.Write(0, 8); bits.Write(0, 1); bits.Write(0, 1); bits.Write(0, 1); // Lossless Y quantizer, no matrix/segmentation.
        bits.Write(0, 1); bits.Align(); // Reduced transform set, no filters or tx-mode bit for lossless.
        bits.Write(0, 1); bits.Align(); // Whole tile group.
        byte[] payload = bits.Bytes.Concat(new byte[] { 0, 11, 1, 22, 23, 2, 33, 34, 35, 44, 45, 46, 47 }).ToArray();
        var (item, sequence) = SyntaxItem(payload, 192, 128);
        Assert.True(OfficeAv1StillFrameReader.TryRead(payload, item, sequence, new OfficeRasterDecodeOptions(), out var frame));
        Assert.True(frame!.CodedLossless);
        Assert.Equal(0, frame.TransformMode);
        Assert.Equal(new[] { 0, 16, 48 }, frame.MiColStarts);
        Assert.Equal(new[] { 0, 16, 32 }, frame.MiRowStarts);
        Assert.Equal(new[] { 1, 2, 3, 4 }, frame.Tiles.Select(t => t.Length));
        Assert.Equal(new[] { 0, 16, 0, 16 }, frame.Tiles.Select(t => t.MiColStart));
        Assert.Equal(new[] { 16, 48, 16, 48 }, frame.Tiles.Select(t => t.MiColEnd));
    }

    [Fact]
    public void SegmentQuantizerControlsLosslessSyntaxAndSignedClamp() {
        var bits = new HeaderWriter();
        bits.Write(0, 3); bits.Write(1, 1); // No screen tools, same render size, uniform single tile.
        bits.Write(0, 8); bits.Write(0, 1); bits.Write(0, 1); bits.Write(1, 1); // Q=0, delta=0, no matrix, segmentation.
        for (int segment = 0; segment < 8; segment++) {
            bits.Write(1, 1); bits.Write(10, 9); // Every segment ALT_Q +10 => not coded lossless.
            bits.Write(1, 1); bits.Write(64, 7); // ALT_LF_Y_V -64 clamps to -63.
            bits.Write(0, 6); // Six disabled features.
        }
        bits.Write(0, 12); bits.Write(0, 3); bits.Write(0, 1); // Luma loop-filter levels 0, sharpness 0, no delta updates.
        bits.Write(1, 1); bits.Write(0, 1); bits.Align(); // Select tx mode, reduced set false.
        byte[] payload = bits.Bytes.Concat(new byte[] { 123 }).ToArray();
        var (item, sequence) = SyntaxItem(payload, 49, 33);
        Assert.True(OfficeAv1StillFrameReader.TryRead(payload, item, sequence, new OfficeRasterDecodeOptions(), out var frame));
        Assert.False(frame!.CodedLossless);
        Assert.Equal(2, frame.TransformMode);
        Assert.Equal(7, frame.LastActiveSegmentId);
        Assert.Equal(-63, frame.SegmentData[0, 1]);
        Assert.All(frame.LosslessSegments, lossless => Assert.False(lossless));
        Assert.Equal(123, payload[Assert.Single(frame.Tiles).Offset]);
    }

    [Fact]
    public void SuperResolutionSeparatesDecodedRenderAndRestorationDimensions() {
        var bits = new HeaderWriter();
        bits.Write(0, 2); bits.Write(1, 1); bits.Write(7, 3); // No screen tools, superres denominator 16.
        bits.Write(1, 1); bits.Write(79, 16); bits.Write(59, 16); // Render metadata 80x60, distinct from 49x33 ispe.
        bits.Write(1, 1); // Uniform single tile.
        bits.Write(0, 8); bits.Write(0, 3); // Q=0, zero DC delta, no matrix/segmentation.
        bits.Write(2, 2); bits.Write(1, 1); bits.Write(1, 1); // Wiener restoration with 256-sample unit.
        bits.Write(0, 1); bits.Align(); // Lossless TX mode implicit; reduced transform set false.
        byte[] payload = bits.Bytes.Concat(new byte[] { 123 }).ToArray();
        var (item, sequence) = SyntaxItem(payload, 49, 33);
        sequence.SuperResolution = sequence.Restoration = true;
        Assert.True(OfficeAv1StillFrameReader.TryRead(payload, item, sequence, new OfficeRasterDecodeOptions(), out var frame));
        Assert.Equal(25, frame!.Width);
        Assert.Equal(49, frame.UpscaledWidth);
        Assert.Equal(80, frame.RenderWidth);
        Assert.Equal(60, frame.RenderHeight);
        Assert.True(frame.CodedLossless);
        Assert.False(frame.AllLossless);
        Assert.Equal(2, frame.RestorationTypes[0]);
        Assert.Equal(256, frame.RestorationUnitSizes[0]);
        Assert.Equal(123, payload[Assert.Single(frame.Tiles).Offset]);
    }

    private static (byte[] Bytes, OfficeAvifImageItem Item, OfficeAv1StillSequence Sequence) ReadFixture(string name, bool alpha) {
        byte[] bytes = File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Avif", name + ".avif"));
        var options = new OfficeRasterDecodeOptions();
        Assert.True(OfficeAvifContainerReader.TryRead(bytes, options, out var container));
        OfficeAvifImageItem item = alpha ? container!.Alpha! : container!.Color;
        Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes, item, options, out var sequence));
        return (bytes, item, sequence!);
    }

    private static (OfficeAvifImageItem Item, OfficeAv1StillSequence Sequence) SyntaxItem(byte[] payload, int width, int height) =>
        (new OfficeAvifImageItem(1, width, height, 0, payload.Length, new byte[] { 129, 0, 28, 0 }, true, null),
        new OfficeAv1StillSequence { MaximumWidth = width, MaximumHeight = height, Monochrome = true, FrameLength = payload.Length });

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
