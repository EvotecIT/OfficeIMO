namespace OfficeIMO.Email.AddressBook;

/// <summary>One discovered OAB component and its non-content metadata.</summary>
public sealed class OfflineAddressBookFileInfo {
    internal OfflineAddressBookFileInfo(string path, string name, long length, uint version,
        OfflineAddressBookFormat format) {
        Path = path;
        Name = name;
        Length = length;
        Version = version;
        Format = format;
    }

    /// <summary>Source path, or caller-supplied stream name.</summary>
    public string Path { get; }
    /// <summary>File name.</summary>
    public string Name { get; }
    /// <summary>Component length.</summary>
    public long Length { get; }
    /// <summary>Little-endian component version marker.</summary>
    public uint Version { get; }
    /// <summary>Recognized component role.</summary>
    public OfflineAddressBookFormat Format { get; }
    /// <summary>Whether this component can supply address entries.</summary>
    public bool CanEnumerateEntries => Format == OfflineAddressBookFormat.Version4FullDetails;
}
