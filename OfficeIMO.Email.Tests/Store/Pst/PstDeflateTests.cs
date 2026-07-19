namespace OfficeIMO.Email.Store.Tests;

public sealed class PstDeflateTests {
    [Fact]
    public void DecodesRawRfc1951PayloadsWithoutAZlibWrapper() {
        byte[] rawDeflate = { 0xCB, 0x48, 0xCD, 0xC9, 0xC9, 0x07, 0x00 };

        byte[] decoded = PstDeflate.Decode(rawDeflate, 5);

        Assert.Equal("hello", Encoding.ASCII.GetString(decoded));
    }

    [Fact]
    public void RejectsDeclaredLengthsThatDoNotMatchTheDeflatePayload() {
        byte[] rawDeflate = { 0xCB, 0x48, 0xCD, 0xC9, 0xC9, 0x07, 0x00 };

        Assert.Throws<InvalidDataException>(() => PstDeflate.Decode(rawDeflate, 4));
        Assert.Throws<InvalidDataException>(() => PstDeflate.Decode(rawDeflate, 6));
    }

    [Fact]
    public void DecodesZlibWrappedPayloadsAndValidatesAdler32() {
        byte[] zlib = { 0x78, 0x9C, 0xCB, 0x48, 0xCD, 0xC9, 0xC9, 0x07, 0x00,
            0x06, 0x2C, 0x02, 0x15 };

        byte[] decoded = PstDeflate.Decode(zlib, 5);

        Assert.Equal("hello", Encoding.ASCII.GetString(decoded));
        zlib[zlib.Length - 1] ^= 0x01;
        Assert.Throws<InvalidDataException>(() => PstDeflate.Decode(zlib, 5));
    }
}
