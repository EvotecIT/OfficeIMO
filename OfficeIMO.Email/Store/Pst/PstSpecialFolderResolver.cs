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
        byte[]? providerUid = GetBinary(storeProperties, MapiKnownProperties.PidTag.RecordKey);
        if (knownFolderNids.Contains(RootFolderNid)) {
            _kinds.Add(RootFolderNid, EmailStoreSpecialFolderKind.Root);
        }

        AddProperties(storeProperties, providerUid, knownFolderNids,
            (MapiKnownProperties.PidTag.IpmSubTreeEntryId, EmailStoreSpecialFolderKind.IpmSubtree),
            (MapiKnownProperties.PidTag.IpmInboxEntryId, EmailStoreSpecialFolderKind.Inbox),
            (MapiKnownProperties.PidTag.IpmOutboxEntryId, EmailStoreSpecialFolderKind.Outbox),
            (MapiKnownProperties.PidTag.IpmWastebasketEntryId, EmailStoreSpecialFolderKind.DeletedItems),
            (MapiKnownProperties.PidTag.IpmSentMailEntryId, EmailStoreSpecialFolderKind.SentItems),
            (MapiKnownProperties.PidTag.ViewsEntryId, EmailStoreSpecialFolderKind.PersonalViews),
            (MapiKnownProperties.PidTag.CommonViewsEntryId, EmailStoreSpecialFolderKind.CommonViews),
            (MapiKnownProperties.PidTag.FinderEntryId, EmailStoreSpecialFolderKind.SearchRoot));

        AddDefaultFolderProperties(rootProperties, providerUid, knownFolderNids);
        AddDefaultFolderProperties(inboxProperties, providerUid, knownFolderNids);
    }

    internal EmailStoreSpecialFolderKind Resolve(uint nid) =>
        _kinds.TryGetValue(nid, out EmailStoreSpecialFolderKind kind)
            ? kind
            : EmailStoreSpecialFolderKind.Unknown;

    internal static bool TryGetStoreFolderNid(IEnumerable<MapiProperty> storeProperties,
        MapiPropertyKey<byte[]> key, out uint nid) {
        if (storeProperties == null) throw new ArgumentNullException(nameof(storeProperties));
        return TryGetFolderNid(storeProperties, key,
            GetBinary(storeProperties, MapiKnownProperties.PidTag.RecordKey), out nid);
    }

    internal static bool TryGetFolderNid(IEnumerable<MapiProperty> properties,
        MapiPropertyKey<byte[]> key, byte[]? providerUid, out uint nid) {
        if (properties == null) throw new ArgumentNullException(nameof(properties));
        byte[]? entryId = GetBinary(properties, key);
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
            (MapiKnownProperties.PidTag.IpmAppointmentEntryId, EmailStoreSpecialFolderKind.Calendar),
            (MapiKnownProperties.PidTag.IpmContactEntryId, EmailStoreSpecialFolderKind.Contacts),
            (MapiKnownProperties.PidTag.IpmJournalEntryId, EmailStoreSpecialFolderKind.Journal),
            (MapiKnownProperties.PidTag.IpmNoteEntryId, EmailStoreSpecialFolderKind.Notes),
            (MapiKnownProperties.PidTag.IpmTaskEntryId, EmailStoreSpecialFolderKind.Tasks),
            (MapiKnownProperties.PidTag.IpmDraftsEntryId, EmailStoreSpecialFolderKind.Drafts));
    }

    private void AddProperties(IEnumerable<MapiProperty>? properties,
        byte[]? providerUid, ISet<uint> knownFolderNids,
        params (MapiPropertyKey<byte[]> Key, EmailStoreSpecialFolderKind Kind)[] mappings) {
        if (properties == null) return;
        foreach ((MapiPropertyKey<byte[]> key, EmailStoreSpecialFolderKind kind) in mappings) {
            if (TryGetFolderNid(properties, key, providerUid, out uint nid) &&
                knownFolderNids.Contains(nid) && !_kinds.ContainsKey(nid)) {
                _kinds.Add(nid, kind);
            }
        }
    }

    private static byte[]? GetBinary(IEnumerable<MapiProperty> properties, MapiPropertyKey<byte[]> key) {
        MapiProperty? property = properties.GetMapiProperty(key);
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
