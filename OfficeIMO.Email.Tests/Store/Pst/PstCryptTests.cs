namespace OfficeIMO.Email.Store.Tests;

public sealed class PstCryptTests {
    [Fact]
    public void DecodesTheDocumentedPermutativeEncoding() {
        byte[] encoded = { 65, 54, 19, 98, 168, 33, 110, 187 };

        PstCrypt.Decode(encoded, 1);

        Assert.Equal(new byte[] { 0, 1, 2, 3, 4, 5, 6, 7 }, encoded);
    }

    [Fact]
    public void CyclicEncodingIsSymmetricForTheBlockBid() {
        byte[] original = Encoding.UTF8.GetBytes("OfficeIMO cyclic PST block");
        byte[] transformed = original.ToArray();
        const ulong bid = 0x123456789ABCDEF0;

        PstCrypt.Decode(transformed, 2, bid);
        Assert.NotEqual(original, transformed);
        PstCrypt.Decode(transformed, 2, bid);

        Assert.Equal(original, transformed);
    }

    [Fact]
    public void CyclicEncodingDependsOnTheAssociatedBlockBid() {
        byte[] first = Enumerable.Range(0, 64).Select(value => (byte)value).ToArray();
        byte[] second = first.ToArray();

        PstCrypt.Decode(first, 2, 0x100);
        PstCrypt.Decode(second, 2, 0x104);

        Assert.NotEqual(first, second);
    }
}
