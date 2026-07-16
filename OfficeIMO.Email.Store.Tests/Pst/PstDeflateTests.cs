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
}
