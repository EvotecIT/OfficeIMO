using OfficeIMO.Email;

namespace OfficeIMO.Email.AddressBook;

/// <summary>Metadata and file-defined schema for one OAB v4 Full Details address list.</summary>
public sealed class OfflineAddressBookListInfo {
    internal OfflineAddressBookListInfo(
        string id,
        int index,
        string sourcePath,
        long sourceLength,
        uint serial,
        long declaredEntryCount,
        IReadOnlyList<OfflineAddressBookPropertyDefinition> headerPropertyDefinitions,
        IReadOnlyList<OfflineAddressBookPropertyDefinition> entryPropertyDefinitions,
        IReadOnlyList<MapiProperty> headerProperties,
        long entriesOffset) {
        Id = id;
        Index = index;
        SourcePath = sourcePath;
        SourceLength = sourceLength;
        Serial = serial;
        DeclaredEntryCount = declaredEntryCount;
        HeaderPropertyDefinitions = headerPropertyDefinitions;
        EntryPropertyDefinitions = entryPropertyDefinitions;
        HeaderProperties = headerProperties;
        EntriesOffset = entriesOffset;
        Name = OabPropertyValues.String(headerProperties, MapiKnownProperties.PidTag.OfflineAddressBookName);
        DistinguishedName = OabPropertyValues.String(headerProperties,
            MapiKnownProperties.PidTag.OfflineAddressBookDistinguishedName);
        Sequence = OabPropertyValues.UInt32(headerProperties, MapiKnownProperties.PidTag.OfflineAddressBookSequence);
        ContainerGuidText = OabPropertyValues.String(headerProperties,
            MapiKnownProperties.PidTag.OfflineAddressBookContainerGuid);
        HierarchicalRootDepartment = OabPropertyValues.String(headerProperties,
            MapiKnownProperties.PidTag.AddressBookHierarchicalRootDepartment);
        if (Guid.TryParse(ContainerGuidText, out Guid parsed)) ContainerGuid = parsed;
    }

    /// <summary>Stable list identifier within the current session.</summary>
    public string Id { get; }
    /// <summary>Zero-based list index.</summary>
    public int Index { get; }
    /// <summary>Source component path or stream name.</summary>
    public string SourcePath { get; }
    /// <summary>Source component length.</summary>
    public long SourceLength { get; }
    /// <summary>Header CRC/serial value.</summary>
    public uint Serial { get; }
    /// <summary>Declared number of address-book-object records.</summary>
    public long DeclaredEntryCount { get; }
    /// <summary>Address-list display name.</summary>
    public string? Name { get; }
    /// <summary>Address-list X500 distinguished name.</summary>
    public string? DistinguishedName { get; }
    /// <summary>OAB generation sequence.</summary>
    public uint? Sequence { get; }
    /// <summary>Container GUID text as encoded in the file.</summary>
    public string? ContainerGuidText { get; }
    /// <summary>Parsed container GUID when valid.</summary>
    public Guid? ContainerGuid { get; }
    /// <summary>Root departmental-group distinguished name when supplied.</summary>
    public string? HierarchicalRootDepartment { get; }
    /// <summary>Header-record schema.</summary>
    public IReadOnlyList<OfflineAddressBookPropertyDefinition> HeaderPropertyDefinitions { get; }
    /// <summary>Address-entry schema.</summary>
    public IReadOnlyList<OfflineAddressBookPropertyDefinition> EntryPropertyDefinitions { get; }
    /// <summary>Decoded header-record values with original property encodings when requested.</summary>
    public IReadOnlyList<MapiProperty> HeaderProperties { get; }

    internal long EntriesOffset { get; }
}
