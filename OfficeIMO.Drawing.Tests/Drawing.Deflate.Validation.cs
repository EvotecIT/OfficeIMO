using OfficeIMO.Core.Internal;
using System.IO.Compression;
using Xunit;

namespace OfficeIMO.Tests;

public sealed class DeflateStreamValidationTests {
    [Fact]
    public void FifteenBitLiteralAndEndCodesPreserveTailAndLimitValidation() {
        // Complete comb-shaped literal tree with 15-bit FF and end codes.
        byte[] payload = Convert.FromBase64String(
            "BeABkCRJkiRJIrGoeWT17AEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAMA//P/+/w==");
        using var input = new MemoryStream(payload);
        using var deflate = new DeflateStream(input, CompressionMode.Decompress);
        Assert.Equal(255, deflate.ReadByte());
        Assert.Equal(-1, deflate.ReadByte());
        Assert.Equal(new byte[] { 255 }, DecodeZlib(payload, new byte[] { 255 }));

        Assert.True(OfficeDeflateStreamValidator.TryValidateExact(payload, 0, payload.Length, 1));
        Assert.False(OfficeDeflateStreamValidator.TryValidateExact(payload, 0, payload.Length, 0, out bool limitExceeded));
        Assert.True(limitExceeded);
        for (int count = 0; count < payload.Length; count++) {
            Assert.False(OfficeDeflateStreamValidator.TryValidateExact(payload, 0, count, 1));
        }
        Array.Resize(ref payload, payload.Length + 1);
        Assert.False(OfficeDeflateStreamValidator.TryValidateExact(payload, 0, payload.Length, 1));
    }

    [Fact]
    public void MixedFixedAndStoredBlocksRetainExactPayloadAndOutputBoundaries() {
        // Non-final fixed block containing A, followed by a final stored BC.
        byte[] payload = { 0x72, 0x04, 0x04, 0x02, 0x00, 0xFD, 0xFF, 0x42, 0x43 };
        using var input = new MemoryStream(payload);
        using var deflate = new DeflateStream(input, CompressionMode.Decompress);
        using var decoded = new MemoryStream();
        deflate.CopyTo(decoded);
        Assert.Equal(new byte[] { 65, 66, 67 }, decoded.ToArray());
        Assert.Equal(new byte[] { 65, 66, 67 }, DecodeZlib(payload, new byte[] { 65, 66, 67 }));

        byte[] surrounded = new byte[payload.Length + 4];
        payload.CopyTo(surrounded, 2);
        Assert.True(OfficeDeflateStreamValidator.TryValidateExact(surrounded, 2, payload.Length, 3));
        Assert.False(OfficeDeflateStreamValidator.TryValidateExact(surrounded, 2, payload.Length + 1, 3));
        for (int count = 0; count < payload.Length; count++) {
            Assert.False(OfficeDeflateStreamValidator.TryValidateExact(surrounded, 2, count, 3));
        }
        Assert.False(OfficeDeflateStreamValidator.TryValidateExact(surrounded, 2, payload.Length, 2, out bool limitExceeded));
        Assert.True(limitExceeded);
    }

    [Theory]
    [InlineData(CompressionLevel.NoCompression)]
    [InlineData(CompressionLevel.Fastest)]
    [InlineData(CompressionLevel.Optimal)]
    public void NativeDeflatePayloadRequiresAnExactEndingAndDeclaredExpansion(CompressionLevel level) {
        var pixels = new byte[65537];
        new Random(1951).NextBytes(pixels);
        for (int index = 0; index < pixels.Length; index++) {
            if (index % 97 < 80) pixels[index] = (byte)(index % 13);
        }
        using var output = new MemoryStream();
        using (var deflate = new DeflateStream(output, level, leaveOpen: true)) {
            deflate.Write(pixels, 0, pixels.Length);
        }
        byte[] payload = output.ToArray();

        Assert.Equal(pixels, DecodeZlib(payload, pixels));

        Assert.True(OfficeDeflateStreamValidator.TryValidateExact(payload, 0, payload.Length, pixels.Length));
        Assert.False(OfficeDeflateStreamValidator.TryValidateExact(payload, 0, payload.Length, pixels.Length - 1, out bool limitExceeded));
        Assert.True(limitExceeded);
        Array.Resize(ref payload, payload.Length + 1);
        Assert.False(OfficeDeflateStreamValidator.TryValidateExact(payload, 0, payload.Length, pixels.Length));
    }

    [Theory]
    [InlineData("eJwFwAEIAAAAACAAAAAAAAAAAAAAAAAAAAAABAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAAEQBCAEI=")]
    [InlineData("eJwFwAEEAAAAQAAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgBAAQgBC")]
    public void ExactZlibDecodeRejectsIncompleteCodeLengthAndLiteralTrees(string encoded) {
        // Independently constructed payloads with an unused canonical code.
        // The advertised output and checksum are valid for the encoded A.
        byte[] zlib = Convert.FromBase64String(encoded);
        Assert.Throws<InvalidDataException>(() =>
            OfficeZlibCodec.Decompress(zlib, maximumOutputBytes: 1, expectedOutputBytes: 1));
    }

    [Fact]
    public void ExactZlibDecodeRequiresTheDeclaredLengthAndChecksum() {
        byte[] zlib = OfficeZlibCodec.Compress(new byte[] { 65, 66, 67 });
        Assert.Throws<InvalidDataException>(() =>
            OfficeZlibCodec.Decompress(zlib, maximumOutputBytes: 4, expectedOutputBytes: 4));
        zlib[zlib.Length - 1] ^= 1;
        Assert.Throws<InvalidDataException>(() =>
            OfficeZlibCodec.Decompress(zlib, maximumOutputBytes: 3, expectedOutputBytes: 3));
    }

    [Theory]
    [InlineData("eJwF3gEEAAAAABAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAEAAAABAEIAQg==")]
    [InlineData("eJwF3wEEAAAAABAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAEAAAACAEIAQg==")]
    public void ExactZlibDecodePreservesRejectionOfReservedDistanceAlphabetSizes(string encoded) {
        // Native-qualified complete trees: 31/32 declared distance symbols,
        // including unused reserved symbols, encode A with a matching checksum.
        byte[] zlib = Convert.FromBase64String(encoded);
        Assert.Throws<InvalidDataException>(() => OfficeZlibCodec.Decompress(zlib, 1, 1));
        Assert.False(OfficeZlibCodec.TryValidateExact(zlib, 0, zlib.Length, 1));
    }

    [Fact]
    public void ExactZlibDecodeAllowsThirtyDeclaredDistanceSymbolsWhenReservedSymbolsAreAbsent() {
        byte[] zlib = Convert.FromBase64String(
            "eJwF3QEEAAAAABAAAAAAAAAAAAEAAAAAAAAAAAAAAAAAAAAAAAAAAAAAgAEAAIAAQgBC");
        Assert.Equal(new byte[] { 65 }, OfficeZlibCodec.Decompress(zlib, 1, 1));
        Assert.True(OfficeZlibCodec.TryValidateExact(zlib, 0, zlib.Length, 1));
    }

    private static byte[] DecodeZlib(byte[] rawDeflate, byte[] expected) {
        var zlib = new byte[rawDeflate.Length + 6];
        zlib[0] = 0x78;
        zlib[1] = 0x9C;
        Buffer.BlockCopy(rawDeflate, 0, zlib, 2, rawDeflate.Length);
        uint checksum = OfficeZlibCodec.Adler32(expected);
        for (int index = 0; index < 4; index++) {
            zlib[zlib.Length - 4 + index] = (byte)(checksum >> (24 - index * 8));
        }
        return OfficeZlibCodec.Decompress(zlib, expected.Length, expected.Length);
    }
}
