namespace OfficeIMO.IWork;

/// <summary>One preserved file entry from an iWork package or bundle.</summary>
public sealed class IWorkPackageEntry {
    private readonly byte[] _bytes;

    internal IWorkPackageEntry(string path, byte[] bytes) {
        Path = path;
        _bytes = bytes;
    }

    /// <summary>Gets the normalized package-relative path.</summary>
    public string Path { get; }
    /// <summary>Gets the uncompressed entry length.</summary>
    public int Length => _bytes.Length;
    /// <summary>Returns a defensive copy of the preserved entry bytes.</summary>
    public byte[] GetBytes() => (byte[])_bytes.Clone();
    internal byte[] Bytes => _bytes;
}
