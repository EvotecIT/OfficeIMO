using System.Text;
using OfficeIMO.Drawing;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class PngCrc32Tests {
    [Fact]
    public void StandardCheckVectorAndEmptyInputMatchCrc32() {
        byte[] check = Encoding.ASCII.GetBytes("123456789");
        Assert.Equal(0xCBF43926U, OfficePngCrc32.Compute(check, 0, check.Length));
        Assert.Equal(0U, OfficePngCrc32.Compute(Array.Empty<byte>(), 0, 0));
    }

    [Theory]
    [InlineData(0)]
    [InlineData(1)]
    [InlineData(7)]
    [InlineData(8)]
    [InlineData(9)]
    [InlineData(15)]
    [InlineData(16)]
    [InlineData(17)]
    [InlineData(31)]
    [InlineData(64)]
    [InlineData(255)]
    [InlineData(256)]
    [InlineData(4095)]
    [InlineData(4096)]
    [InlineData(4097)]
    [InlineData(65537)]
    public void OffsetSlicesAndExistingStatesMatchBitwiseReference(int count) {
        var bytes = new byte[count + 32];
        new Random(1951 + count).NextBytes(bytes);
        foreach (int offset in new[] { 0, 1, 3, 7, 8, 15 }) {
            foreach (uint initial in new[] { 0U, 1U, uint.MaxValue, 0xDEADBEEFU }) {
                Assert.Equal(AppendBitwise(initial, bytes, offset, count),
                    OfficePngCrc32.Append(initial, bytes, offset, count));
            }
        }
    }

    [Fact]
    public void IncrementalChunksMatchOneContinuousChecksum() {
        var bytes = new byte[8193];
        new Random(1951).NextBytes(bytes);
        uint expected = AppendBitwise(OfficePngCrc32.Begin(), bytes, 0, bytes.Length);
        foreach (int chunkSize in new[] { 1, 3, 7, 8, 9, 15, 16, 17, 4095, 4096, 4097 }) {
            uint actual = OfficePngCrc32.Begin();
            for (int offset = 0; offset < bytes.Length; offset += chunkSize) {
                actual = OfficePngCrc32.Append(actual, bytes, offset, Math.Min(chunkSize, bytes.Length - offset));
            }
            Assert.Equal(expected, actual);
            Assert.Equal(expected ^ uint.MaxValue, OfficePngCrc32.Complete(actual));
        }
    }

    private static uint AppendBitwise(uint crc, byte[] bytes, int offset, int count) {
        for (int index = offset; index < offset + count; index++) {
            crc ^= bytes[index];
            for (int bit = 0; bit < 8; bit++) {
                crc = (crc >> 1) ^ ((crc & 1) == 0 ? 0U : 0xEDB88320U);
            }
        }
        return crc;
    }
}
