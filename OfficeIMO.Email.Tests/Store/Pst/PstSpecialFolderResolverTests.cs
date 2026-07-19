using OfficeIMO.Email;

namespace OfficeIMO.Email.Store.Tests;

public sealed class PstSpecialFolderResolverTests {
    [Fact]
    public void ResolvesProviderMatchedFolderEntryIds() {
        byte[] providerUid = Enumerable.Range(1, 16).Select(value => (byte)value).ToArray();
        var storeProperties = new[] {
            BinaryProperty(0x0FF9, providerUid),
            BinaryProperty(0x35E0, CreateEntryId(providerUid, 0x8022))
        };

        var resolver = new PstSpecialFolderResolver(
            storeProperties, null, null, new uint[] { 0x8022 });

        Assert.Equal(EmailStoreSpecialFolderKind.IpmSubtree, resolver.Resolve(0x8022));
    }

    [Fact]
    public void RejectsForeignMalformedAndNonFolderEntryIds() {
        byte[] providerUid = Enumerable.Range(1, 16).Select(value => (byte)value).ToArray();
        byte[] foreignUid = providerUid.Select(value => (byte)(value + 1)).ToArray();
        var foreign = new[] { BinaryProperty(0x35E0, CreateEntryId(foreignUid, 0x8022)) };
        var nonFolder = new[] { BinaryProperty(0x35E0, CreateEntryId(providerUid, 0x8004)) };
        var malformed = new[] { BinaryProperty(0x35E0, new byte[23]) };

        Assert.False(PstSpecialFolderResolver.TryGetFolderNid(
            foreign, MapiKnownProperties.PidTag.IpmSubTreeEntryId, providerUid, out _));
        Assert.False(PstSpecialFolderResolver.TryGetFolderNid(
            nonFolder, MapiKnownProperties.PidTag.IpmSubTreeEntryId, providerUid, out _));
        Assert.False(PstSpecialFolderResolver.TryGetFolderNid(
            malformed, MapiKnownProperties.PidTag.IpmSubTreeEntryId, providerUid, out _));
    }

    private static MapiProperty BinaryProperty(ushort id, byte[] value) =>
        new MapiProperty(id, MapiPropertyType.Binary, value) { RawData = value };

    private static byte[] CreateEntryId(byte[] providerUid, uint nid) {
        var result = new byte[24];
        Buffer.BlockCopy(providerUid, 0, result, 4, providerUid.Length);
        result[20] = (byte)nid;
        result[21] = (byte)(nid >> 8);
        result[22] = (byte)(nid >> 16);
        result[23] = (byte)(nid >> 24);
        return result;
    }
}
