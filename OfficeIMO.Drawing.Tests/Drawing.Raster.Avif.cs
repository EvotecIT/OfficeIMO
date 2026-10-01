using System.Threading;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

/// <summary>Independent libavif still items exercise container safety before the pixel decoder consumes a tile.</summary>
public sealed class DrawingAvifTests {
    [Theory]
    [InlineData("avif-opaque", false)]
    [InlineData("avif-alpha", true)]
    public void IndependentStillItemsLocateColorAndAuxiliaryAlpha(string name, bool hasAlpha) {
        byte[] bytes = Fixture(name);
        Assert.True(OfficeAvifContainerReader.TryRead(bytes, new OfficeRasterDecodeOptions(), out var container));
        Assert.NotNull(container);
        Assert.Equal(49, container.Color.Width);
        Assert.Equal(33, container.Color.Height);
        Assert.Equal(hasAlpha, container.Alpha != null);
        Assert.False(container.Color.Monochrome);
        Assert.NotNull(container.Color.ColorDescription);
        Assert.Equal(1, container.Color.ColorDescription.Primaries);
        Assert.Equal(13, container.Color.ColorDescription.Transfer);
        Assert.Equal(6, container.Color.ColorDescription.Matrix);
        Assert.True(container.Color.ColorDescription.FullRange);
        byte[] opaque = Fixture("avif-opaque");
        Assert.True(OfficeAvifContainerReader.TryRead(opaque, new OfficeRasterDecodeOptions(), out var opaqueContainer));
        Assert.Equal(Payload(opaque, opaqueContainer!.Color), Payload(bytes, container.Color));
        if (hasAlpha) {
            Assert.True(container.Alpha!.Monochrome);
            Assert.Equal(container.Color.Width, container.Alpha.Width);
            Assert.Equal(container.Color.Height, container.Alpha.Height);
            Assert.Null(container.Alpha.ColorDescription);
            Assert.NotEqual(Payload(bytes, container.Color), Payload(bytes, container.Alpha));
        }
    }

    [Theory]
    [InlineData("avif-opaque")]
    [InlineData("avif-alpha")]
    public void TruncatedContainerNeverBorrowsBytesFromMissingMediaData(string name) {
        byte[] bytes = Fixture(name);
        for (int length = 0; length < bytes.Length; length++) {
            var truncated = new byte[length];
            Buffer.BlockCopy(bytes, 0, truncated, 0, length);
            Assert.False(OfficeAvifContainerReader.TryRead(truncated, new OfficeRasterDecodeOptions(), out _));
        }
    }

    [Fact]
    public void ContainerHonorsEncodedPixelAndCancellationLimits() {
        byte[] bytes = Fixture("avif-alpha");
        Assert.False(OfficeAvifContainerReader.TryRead(bytes, new OfficeRasterDecodeOptions { MaximumEncodedBytes = bytes.Length - 1 }, out _));
        Assert.False(OfficeAvifContainerReader.TryRead(bytes, new OfficeRasterDecodeOptions { MaximumDecodedPixels = 49 * 33 - 1 }, out _));
        Assert.Throws<OperationCanceledException>(() => OfficeAvifContainerReader.TryRead(bytes,
            new OfficeRasterDecodeOptions { CancellationToken = new CancellationToken(true) }, out _));
    }

    [Theory]
    [InlineData("avif-opaque")]
    [InlineData("avif-alpha")]
    public void IndependentSequenceHeadersMatchConfigurationAndRetainWholeFrame(string name) {
        byte[] bytes = Fixture(name);
        var options = new OfficeRasterDecodeOptions();
        Assert.True(OfficeAvifContainerReader.TryRead(bytes, options, out var container));
        var items = container!.Alpha == null ? new[] { container.Color } : new[] { container.Color, container.Alpha };
        foreach (var item in items) {
            Assert.True(OfficeAv1StillSequenceReader.TryRead(bytes, item!, options, out var sequence));
            Assert.NotNull(sequence);
            Assert.Equal(49, sequence.MaximumWidth);
            Assert.Equal(33, sequence.MaximumHeight);
            Assert.Equal(item!.Monochrome, sequence.Monochrome);
            Assert.True(sequence.Color.FullRange);
            Assert.InRange(sequence.FrameOffset, item.Offset, item.Offset + item.Length - 1);
            Assert.Equal(item.Offset + item.Length, sequence.FrameOffset + sequence.FrameLength);
        }
    }

    [Fact]
    public void SequenceSizeCannotEscapeItemAndConfigurationMustMatchBitstream() {
        byte[] original = Fixture("avif-opaque");
        var options = new OfficeRasterDecodeOptions();
        Assert.True(OfficeAvifContainerReader.TryRead(original, options, out var container));
        OfficeAvifImageItem item = container!.Color;
        byte[] bytes = (byte[])original.Clone();
        bytes[item.Offset + 3] = 127; // Sequence OBU's LEB128 size extends past the bounded item.
        Assert.False(OfficeAv1StillSequenceReader.TryRead(bytes, item, options, out _));
        bytes = (byte[])original.Clone();
        bytes[item.Offset + 4] ^= 1; // Change declared sequence level without changing av1C.
        Assert.False(OfficeAv1StillSequenceReader.TryRead(bytes, item, options, out _));
        bytes = (byte[])original.Clone();
        bytes[item.Offset] |= 128; // Forbidden OBU bit.
        Assert.False(OfficeAv1StillSequenceReader.TryRead(bytes, item, options, out _));
        Assert.Throws<OperationCanceledException>(() => OfficeAv1StillSequenceReader.TryRead(original, item,
            new OfficeRasterDecodeOptions { CancellationToken = new CancellationToken(true) }, out _));
    }

    [Fact]
    public void LastFrameMayUseRemainingItemBytesWithoutExplicitObuSize() {
        byte[] original = Fixture("avif-opaque");
        var options = new OfficeRasterDecodeOptions();
        Assert.True(OfficeAvifContainerReader.TryRead(original, options, out var container));
        OfficeAvifImageItem item = container!.Color;
        Assert.True(OfficeAv1StillSequenceReader.TryRead(original, item, options, out var sequence));
        int sizeByte = sequence!.FrameOffset - 1;
        byte[] payload = new byte[item.Length - 1];
        Buffer.BlockCopy(original, item.Offset, payload, 0, sizeByte - item.Offset);
        payload[sizeByte - item.Offset - 1] &= unchecked((byte)~2); // Clear frame OBU size flag.
        Buffer.BlockCopy(original, sizeByte + 1, payload, sizeByte - item.Offset, original.Length - sizeByte - 1);
        var remaining = new OfficeAvifImageItem(item.Id, item.Width, item.Height, 0, payload.Length,
            item.Configuration, item.IsAlpha, item.ColorDescription);
        Assert.True(OfficeAv1StillSequenceReader.TryRead(payload, remaining, options, out var noSizeSequence));
        Assert.Equal(sequence.FrameLength, noSizeSequence!.FrameLength);
        Assert.Equal(payload.Length, noSizeSequence.FrameOffset + noSizeSequence.FrameLength);
    }

    [Fact]
    public void TruncatedHeaderCannotConsumeFollowingFrameBytes() {
        byte[] original = Fixture("avif-opaque");
        var options = new OfficeRasterDecodeOptions();
        Assert.True(OfficeAvifContainerReader.TryRead(original, options, out var container));
        for (int length = 0; length < 9; length++) {
            byte[] bytes = (byte[])original.Clone();
            bytes[container!.Color.Offset + 3] = (byte)length;
            Assert.False(OfficeAv1StillSequenceReader.TryRead(bytes, container.Color, options, out _));
        }
    }

    [Theory]
    [InlineData("18157081a202020108", false)]
    [InlineData("18157081a602020140", true)]
    public void IdentityMatrixCannotDescribeSubsampledColorOrMonochrome(string sequenceHex, bool monochrome) {
        byte[] bytes = Fixture("avif-opaque");
        var options = new OfficeRasterDecodeOptions();
        Assert.True(OfficeAvifContainerReader.TryRead(bytes, options, out var container));
        OfficeAvifImageItem color = container!.Color;
        byte[] sequence = ConvertHex(sequenceHex);
        Buffer.BlockCopy(sequence, 0, bytes, color.Offset + 4, sequence.Length);
        byte[] configuration = (byte[])color.Configuration.Clone();
        if (monochrome) configuration[2] |= 16;
        var item = new OfficeAvifImageItem(color.Id, color.Width, color.Height, color.Offset, color.Length,
            configuration, false, null);
        Assert.False(OfficeAv1StillSequenceReader.TryRead(bytes, item, options, out _));
    }

    [Fact]
    public void ItemExtentCannotPointAtMetadataOrOutsideItsMediaBox() {
        byte[] original = Fixture("avif-opaque");
        int location = Tag(original, "iloc") + 4;
        foreach (uint offset in new uint[] { 0, 32, 267, uint.MaxValue }) {
            byte[] bytes = (byte[])original.Clone();
            Write32(bytes, location + 14, offset);
            Assert.False(OfficeAvifContainerReader.TryRead(bytes, new OfficeRasterDecodeOptions(), out _));
        }
    }

    [Fact]
    public void SelectedUnknownEssentialPropertyAndMalformedAlphaAreNotDiscarded() {
        byte[] opaque = Fixture("avif-opaque");
        byte[] bytes = (byte[])opaque.Clone();
        Encoding.ASCII.GetBytes("zzzz").CopyTo(bytes, Tag(bytes, "av1C"));
        Assert.False(OfficeAvifContainerReader.TryRead(bytes, new OfficeRasterDecodeOptions(), out _));

        bytes = Fixture("avif-alpha");
        int auxiliary = Tag(bytes, "auxC");
        bytes[auxiliary + 8] = (byte)'x';
        Assert.False(OfficeAvifContainerReader.TryRead(bytes, new OfficeRasterDecodeOptions(), out _));

        bytes = (byte[])opaque.Clone();
        Encoding.ASCII.GetBytes("zzzz").CopyTo(bytes, Tag(bytes, "pixi"));
        Assert.True(OfficeAvifContainerReader.TryRead(bytes, new OfficeRasterDecodeOptions(), out _)); // Unknown optional description is safe.
    }

    [Theory]
    [InlineData("clap")]
    [InlineData("irot")]
    [InlineData("imir")]
    public void UnimplementedTransformsAreRejectedEvenIfNotMarkedEssential(string transform) {
        byte[] bytes = Fixture("avif-opaque");
        Encoding.ASCII.GetBytes(transform).CopyTo(bytes, Tag(bytes, "pixi"));
        Assert.False(OfficeAvifContainerReader.TryRead(bytes, new OfficeRasterDecodeOptions(), out _));
    }

    private static byte[] Fixture(string name) => File.ReadAllBytes(Path.Combine(AppContext.BaseDirectory, "TestAssets", "Avif", name + ".avif"));
    private static byte[] ConvertHex(string hex) {
        var result = new byte[hex.Length / 2];
        for (int i = 0; i < result.Length; i++) result[i] = Convert.ToByte(hex.Substring(i * 2, 2), 16);
        return result;
    }
    private static byte[] Payload(byte[] bytes, OfficeAvifImageItem item) {
        var payload = new byte[item.Length];
        Buffer.BlockCopy(bytes, item.Offset, payload, 0, payload.Length);
        return payload;
    }
    private static int Tag(byte[] bytes, string type) {
        byte[] tag = Encoding.ASCII.GetBytes(type);
        for (int i = 0; i <= bytes.Length - 4; i++)
            if (bytes[i] == tag[0] && bytes[i + 1] == tag[1] && bytes[i + 2] == tag[2] && bytes[i + 3] == tag[3]) return i;
        throw new InvalidOperationException("Missing fixture box " + type);
    }
    private static void Write32(byte[] bytes, int p, uint value) {
        bytes[p] = (byte)(value >> 24); bytes[p + 1] = (byte)(value >> 16);
        bytes[p + 2] = (byte)(value >> 8); bytes[p + 3] = (byte)value;
    }
}
