using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class JpegRawComponentsTests {
    [Theory]
    [InlineData(6)]
    [InlineData(255)]
    public void SeparateComponentFramesRespectTheFrameCountWithoutInventingRgbaColors(int count) {
        byte[] jpeg = NeutralFrame(count);
        Assert.True(OfficeJpegCodec.TryDecodeColorComponents(jpeg, 0, false,
            out byte[] components, out int width, out int height, out int actualCount));
        Assert.Equal((1, 1, count), (width, height, actualCount));
        Assert.Equal(Enumerable.Repeat((byte)128, count), components);
        Assert.False(OfficeJpegCodec.TryDecode(jpeg, out _));
        Assert.False(OfficeJpegCodec.TryDecodeColorComponents(jpeg, null, false, out _, out _, out _, out _));
    }

    [Fact]
    public void EveryRawOutputChannelCountsTowardTheWorkingSetBudget() {
        long retained = OfficeRasterGuards.MaximumDecodedBytes - 64L * 1024 - 4L * 1024 * 1024;
        Assert.True(OfficeJpegReader.TryInitializeDecodeWorkingSet(retained, 1024, 1024, 1, out _, 4));
        Assert.False(OfficeJpegReader.TryInitializeDecodeWorkingSet(retained, 1024, 1024, 1, out _, 255));
    }

    // A complete one-pixel JPEG with one sequential scan per component, DC=0, AC=EOB.
    private static byte[] NeutralFrame(int count) {
        var bytes = new List<byte> { 255, 216 };
        void Segment(byte marker, IEnumerable<byte> payload) {
            byte[] body = payload.ToArray(); int length = body.Length + 2;
            bytes.AddRange(new byte[] { 255, marker, (byte)(length >> 8), (byte)length }); bytes.AddRange(body);
        }
        Segment(219, new byte[] { 0 }.Concat(Enumerable.Repeat((byte)1, 64)));
        Segment(196, new byte[] { 0, 1 }.Concat(new byte[15]).Concat(new byte[] { 0, 16, 1 }).Concat(new byte[15]).Concat(new byte[] { 0 }));
        Segment(192, new byte[] { 8, 0, 1, 0, 1, (byte)count }.Concat(
            Enumerable.Range(0, count).SelectMany(i => new byte[] { (byte)i, 17, 0 })));
        for (int i = 0; i < count; i++) { Segment(218, new byte[] { 1, (byte)i, 0, 0, 63, 0 }); bytes.Add(0x3F); }
        bytes.AddRange(new byte[] { 255, 217 }); return bytes.ToArray();
    }
}
