#if NET8_0_OR_GREATER
using System.IO.Compression;
using OfficeIMO.Core.Internal;

namespace OfficeIMO.Archives.Tests;

public sealed class ZipDeflateInflationTests {
    [Theory]
    [InlineData(0, 1)]
    [InlineData(1, 1)]
    [InlineData(272, 17)]
    [InlineData(273, 17)]
    [InlineData(4097, 3)]
    [InlineData(32769, 32768)]
    [InlineData(65537, 7)]
    [InlineData(131073, 31)]
    public void BoundedInflationPreservesLogicalOutputAndPublishedPrefixes(int length, int period) {
        byte[] expected = new byte[length];
        for (int index = 0; index < length; index++) {
            int value = index % period;
            expected[index] = (byte)((value * 37) ^ (value >> 8));
        }

        foreach (CompressionLevel level in new[] {
            CompressionLevel.NoCompression, CompressionLevel.Fastest, CompressionLevel.Optimal
        }) {
            byte[] compressed = Compress(expected, level);
            byte[] input = Enumerable.Repeat((byte)0x9d, compressed.Length + 8).ToArray();
            compressed.CopyTo(input, 3);
            byte[] output = Enumerable.Repeat((byte)0xa5, length + 32).ToArray();
            int published = 0;

            Assert.True(OfficeDeflateStreamValidator.TryInflateExact(
                input, 3, compressed.Length, output, length, written => {
                    Assert.InRange(written, published + 1, length);
                    Assert.True(expected.AsSpan(0, written).SequenceEqual(output.AsSpan(0, written)));
                    published = written;
                }));
            Assert.True(expected.AsSpan().SequenceEqual(output.AsSpan(0, length)));
            Assert.All(output.Skip(length), value => Assert.Equal((byte)0xa5, value));

            Assert.False(OfficeDeflateStreamValidator.TryInflateExact(
                input, 3, compressed.Length + 1, output, length, null));
            if (length == 0) continue;
            Array.Fill(output, (byte)0xa5);
            Assert.False(OfficeDeflateStreamValidator.TryInflateExact(
                input, 3, compressed.Length, output, length - 1, null));
            Assert.All(output.Skip(length - 1), value => Assert.Equal((byte)0xa5, value));
        }
    }

    [Fact]
    public void WideMatchCopiesPreserveTheLogicalOutputTail() {
        // Independent raw DEFLATE: a repeated 16-byte pattern, then empty blocks.
        // The following blocks retain lookahead while the last wide match reaches the output boundary.
        byte[] compressed = Convert.FromHexString(
            "6250f5ca9fb2f31eb3866fd1f43d0fd9b419463c1f000000ffff000000ffff000000ffff010000ffff");
        byte[] expected = Enumerable.Range(0, 274).Select(index => (byte)((index % 16) * 37)).ToArray();
        using (var bcl = new DeflateStream(new MemoryStream(compressed), CompressionMode.Decompress)) {
            using var decoded = new MemoryStream();
            bcl.CopyTo(decoded);
            Assert.Equal(expected, decoded.ToArray());
        }
        byte[] output = Enumerable.Repeat((byte)0xa5, expected.Length + 32).ToArray();
        Assert.True(OfficeDeflateStreamValidator.TryInflateExact(
            compressed, 0, compressed.Length, output, expected.Length, null));
        Assert.True(expected.AsSpan().SequenceEqual(output.AsSpan(0, expected.Length)));
        Assert.All(output.Skip(expected.Length), value => Assert.Equal((byte)0xa5, value));
    }

    private static byte[] Compress(byte[] bytes, CompressionLevel level) {
        using var result = new MemoryStream();
        using (var compressor = new DeflateStream(result, level, leaveOpen: true)) {
            compressor.Write(bytes);
        }
        // An unwritten BCL stream can emit no payload; this is the RFC 1951 empty stored block.
        return result.Length == 0 ? new byte[] { 1, 0, 0, 0xff, 0xff } : result.ToArray();
    }
}
#endif
