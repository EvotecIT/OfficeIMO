using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

internal sealed class PstSpecialFolderResolver {
    private const uint RootFolderNid = 0x00000122;
    private readonly Dictionary<uint, EmailStoreSpecialFolderKind> _kinds =
        new Dictionary<uint, EmailStoreSpecialFolderKind>();

    internal PstSpecialFolderResolver(
        IEnumerable<MapiProperty> storeProperties,
        IEnumerable<MapiProperty>? rootProperties,
        IEnumerable<MapiProperty>? inboxProperties,
        IReadOnlyCollection<uint> folderNids) {
        if (storeProperties == null) throw new ArgumentNullException(nameof(storeProperties));
        if (folderNids == null) throw new ArgumentNullException(nameof(folderNids));

        var knownFolderNids = new HashSet<uint>(folderNids);
        byte[]? providerUid = GetBinary(storeProperties, 0x0FF9);
        if (knownFolderNids.Contains(RootFolderNid)) {
            _kinds.Add(RootFolderNid, EmailStoreSpecialFolderKind.Root);
        }

        AddProperties(storeProperties, providerUid, knownFolderNids,
            (0x35E0, EmailStoreSpecialFolderKind.IpmSubtree),
            (0x35E1, EmailStoreSpecialFolderKind.Inbox),
            (0x35E2, EmailStoreSpecialFolderKind.Outbox),
            (0x35E3, EmailStoreSpecialFolderKind.DeletedItems),
            (0x35E4, EmailStoreSpecialFolderKind.SentItems),
            (0x35E5, EmailStoreSpecialFolderKind.PersonalViews),
            (0x35E6, EmailStoreSpecialFolderKind.CommonViews),
            (0x35E7, EmailStoreSpecialFolderKind.SearchRoot));

        AddDefaultFolderProperties(rootProperties, providerUid, knownFolderNids);
        AddDefaultFolderProperties(inboxProperties, providerUid, knownFolderNids);
    }

    internal EmailStoreSpecialFolderKind Resolve(uint nid) =>
        _kinds.TryGetValue(nid, out EmailStoreSpecialFolderKind kind)
            ? kind
            : EmailStoreSpecialFolderKind.Unknown;

    internal static bool TryGetStoreFolderNid(IEnumerable<MapiProperty> storeProperties,
        ushort propertyId, out uint nid) {
        if (storeProperties == null) throw new ArgumentNullException(nameof(storeProperties));
        return TryGetFolderNid(storeProperties, propertyId,
            GetBinary(storeProperties, 0x0FF9), out nid);
    }

    internal static bool TryGetFolderNid(IEnumerable<MapiProperty> properties,
        ushort propertyId, byte[]? providerUid, out uint nid) {
        if (properties == null) throw new ArgumentNullException(nameof(properties));
        byte[]? entryId = GetBinary(properties, propertyId);
        if (entryId == null || entryId.Length != 24 ||
            entryId[0] != 0 || entryId[1] != 0 || entryId[2] != 0 || entryId[3] != 0) {
            nid = 0;
            return false;
        }
        if (providerUid != null &&
            (providerUid.Length != 16 || !EqualsAt(entryId, 4, providerUid))) {
            nid = 0;
            return false;
        }

        nid = PstBinary.UInt32(entryId, 20);
        byte nodeType = (byte)(nid & 0x1F);
        if (nodeType == 0x02 || nodeType == 0x03) return true;
        nid = 0;
        return false;
    }

    private void AddDefaultFolderProperties(IEnumerable<MapiProperty>? properties,
        byte[]? providerUid, ISet<uint> knownFolderNids) {
        AddProperties(properties, providerUid, knownFolderNids,
            (0x36D0, EmailStoreSpecialFolderKind.Calendar),
            (0x36D1, EmailStoreSpecialFolderKind.Contacts),
            (0x36D2, EmailStoreSpecialFolderKind.Journal),
            (0x36D3, EmailStoreSpecialFolderKind.Notes),
            (0x36D4, EmailStoreSpecialFolderKind.Tasks),
            (0x36D7, EmailStoreSpecialFolderKind.Drafts));
    }

    private void AddProperties(IEnumerable<MapiProperty>? properties,
        byte[]? providerUid, ISet<uint> knownFolderNids,
        params (ushort PropertyId, EmailStoreSpecialFolderKind Kind)[] mappings) {
        if (properties == null) return;
        foreach ((ushort propertyId, EmailStoreSpecialFolderKind kind) in mappings) {
            if (TryGetFolderNid(properties, propertyId, providerUid, out uint nid) &&
                knownFolderNids.Contains(nid) && !_kinds.ContainsKey(nid)) {
                _kinds.Add(nid, kind);
            }
        }
    }

    private static byte[]? GetBinary(IEnumerable<MapiProperty> properties, ushort propertyId) {
        MapiProperty? property = properties.FirstOrDefault(item => item.PropertyId == propertyId);
        return property?.Value as byte[] ?? property?.RawData;
    }

    private static bool EqualsAt(byte[] source, int offset, byte[] expected) {
        if (offset < 0 || offset > source.Length - expected.Length) return false;
        for (int index = 0; index < expected.Length; index++) {
            if (source[offset + index] != expected[index]) return false;
        }
        return true;
    }
}
