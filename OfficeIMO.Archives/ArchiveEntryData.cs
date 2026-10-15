namespace OfficeIMO.Archives;

/// <summary>One fully verified selected file. Its content is owned by the caller and independent of the input stream.</summary>
public sealed class ArchiveEntryData {
    internal ArchiveEntryData(string relativePath, byte[] data) { RelativePath = relativePath; Data = data; }
    /// <summary>Portable relative member path, preserving case and using forward slashes.</summary>
    public string RelativePath { get; }
    /// <summary>Complete decoded bytes, validated against the declared length and ZIP CRC-32.</summary>
    public byte[] Data { get; }
    /// <summary>Decoded content length in bytes.</summary>
    public long Length => Data.LongLength;
}
