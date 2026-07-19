using OfficeIMO.Email;

namespace OfficeIMO.Email.Store;

/// <summary>Persistent Outlook search-folder definition message and bounded definition header.</summary>
public sealed class EmailStoreSearchFolderDefinition {
    private const string ExpectedMessageClass = "IPM.Microsoft.Wunderbar.SFInfo";
    private const uint ExpectedDefinitionVersion = 0x04100000;
    private const uint KnownFlags = 0x0000707F;

    internal EmailStoreSearchFolderDefinition(EmailDocument document) {
        MessageClass = document.MessageClass;
        TemplateId = document.Mapi.GetNullableValue(MapiKnownProperties.PidTag.SearchFolderTemplateId);
        SearchFolderId = Copy(document.Mapi.GetValueOrDefault(MapiKnownProperties.PidTag.SearchFolderId));
        Definition = Copy(document.Mapi.GetValueOrDefault(MapiKnownProperties.PidTag.SearchFolderDefinition));
        StorageType = document.Mapi.GetNullableValue(MapiKnownProperties.PidTag.SearchFolderStorageType);
        SynchronizationTag = document.Mapi.GetNullableValue(MapiKnownProperties.PidTag.SearchFolderTag);
        ExtendedFolderFlags = document.Mapi.GetNullableValue(MapiKnownProperties.PidTag.SearchFolderEfpFlags);
        LastUsedAt = MinutesSince1601(document.Mapi.GetNullableValue(MapiKnownProperties.PidTag.SearchFolderLastUsed));
        ExpiresAt = MinutesSince1601(document.Mapi.GetNullableValue(MapiKnownProperties.PidTag.SearchFolderExpiration));
        if (Definition != null && Definition.Length >= 12) {
            DefinitionVersion = ReadUInt32BigEndian(Definition, 0);
            DefinitionFlags = ReadUInt32BigEndian(Definition, 4);
            NumericalSearch = ReadUInt32BigEndian(Definition, 8);
        }
    }

    /// <summary>Actual associated-message class.</summary>
    public string? MessageClass { get; }
    /// <summary>Search template identifier.</summary>
    public int? TemplateId { get; }
    /// <summary>Raw search-folder identifier that correlates with a container's extended folder flags.</summary>
    public byte[]? SearchFolderId { get; }
    /// <summary>Exact template-dependent PidTagSearchFolderDefinition BLOB.</summary>
    public byte[]? Definition { get; }
    /// <summary>Flags controlling the fields present in <see cref="Definition"/>.</summary>
    public int? StorageType { get; }
    /// <summary>Definition/container synchronization tag.</summary>
    public int? SynchronizationTag { get; }
    /// <summary>Extended flags copied to the search-folder container.</summary>
    public int? ExtendedFolderFlags { get; }
    /// <summary>Last-used time decoded from minutes since 1601.</summary>
    public DateTimeOffset? LastUsedAt { get; }
    /// <summary>Expiration time decoded from minutes since 1601.</summary>
    public DateTimeOffset? ExpiresAt { get; }
    /// <summary>Network-order definition version; 0x04100000 is defined.</summary>
    public uint? DefinitionVersion { get; }
    /// <summary>Network-order field-presence and refresh flags.</summary>
    public uint? DefinitionFlags { get; }
    /// <summary>Template-specific size/age number when the numerical flag is set.</summary>
    public uint? NumericalSearch { get; }
    /// <summary>Whether an EntryList of folders is present.</summary>
    public bool HasFolderEntryList => HasFlag(0x00000040);
    /// <summary>Whether a textual folder-name list is present.</summary>
    public bool HasFolderNameList => HasFlag(0x00000020);
    /// <summary>Whether implementation-specific advanced-search bytes are present.</summary>
    public bool HasAdvancedSearch => HasFlag(0x00000010);
    /// <summary>Whether a documented MAPI restriction packet is present.</summary>
    public bool HasRestriction => HasFlag(0x00000008);
    /// <summary>Whether an address list is present.</summary>
    public bool HasAddresses => HasFlag(0x00000004);
    /// <summary>Whether textual search criteria are present.</summary>
    public bool HasTextSearch => HasFlag(0x00000002);
    /// <summary>Whether <see cref="NumericalSearch"/> is active.</summary>
    public bool UsesNumericalSearch => HasFlag(0x00000001);
    /// <summary>Whether the container requests daily refresh.</summary>
    public bool RefreshDaily => HasFlag(0x00004000);
    /// <summary>Whether the container requests weekly refresh.</summary>
    public bool RefreshWeekly => HasFlag(0x00002000);
    /// <summary>Whether the container requests monthly refresh.</summary>
    public bool RefreshMonthly => HasFlag(0x00001000);

    /// <summary>True when the message class and fixed definition header are valid.</summary>
    public bool IsProtocolEnvelopeValid =>
        string.Equals(MessageClass, ExpectedMessageClass, StringComparison.OrdinalIgnoreCase) &&
        TemplateId.HasValue && SearchFolderId != null && Definition != null && Definition.Length >= 12 &&
        DefinitionVersion == ExpectedDefinitionVersion && DefinitionFlags.HasValue &&
        (DefinitionFlags.Value & ~KnownFlags) == 0;

    private bool HasFlag(uint flag) => (DefinitionFlags.GetValueOrDefault() & flag) != 0;
    private static uint ReadUInt32BigEndian(byte[] bytes, int offset) =>
        ((uint)bytes[offset] << 24) | ((uint)bytes[offset + 1] << 16) |
        ((uint)bytes[offset + 2] << 8) | bytes[offset + 3];
    private static byte[]? Copy(byte[]? value) => value == null ? null : (byte[])value.Clone();
    private static DateTimeOffset? MinutesSince1601(int? value) {
        if (!value.HasValue || value.Value < 0) return null;
        try {
            return new DateTimeOffset(1601, 1, 1, 0, 0, 0, TimeSpan.Zero).AddMinutes(value.Value);
        } catch (ArgumentOutOfRangeException) {
            return null;
        }
    }
}
