namespace OfficeIMO.OpenDocument;

public abstract partial class OdfDocument {
    /// <summary>Package entry paths currently stored in the document.</summary>
    public IReadOnlyList<string> PackageEntries => Package.Entries.Select(entry => entry.Name).ToList();

    /// <summary>Normalizes a package reference through the native owner without fetching or resolving filesystem targets.</summary>
    public static string NormalizePackageEntryPath(string reference) {
        if (reference == null) throw new ArgumentNullException(nameof(reference));
        return OdfPackagePath.NormalizeHref(reference);
    }

    /// <summary>Declared media type of an already-loaded package entry. External references are never fetched.</summary>
    public string? GetPackageEntryMediaType(string reference) => Package.GetRequiredEntry(NormalizePackageEntryPath(reference)).MediaType;

    /// <summary>Decoded byte length of an already-loaded package entry, without creating a defensive payload copy.</summary>
    public long GetPackageEntryByteLength(string reference) => Package.GetRequiredEntry(NormalizePackageEntryPath(reference)).GetBytesForSave().LongLength;

    /// <summary>Returns a defensive copy of one already-loaded package entry without resolving external links.</summary>
    public byte[] GetPackageEntryBytes(string path) => GetPackageEntryBytes(path, long.MaxValue);

    /// <summary>Returns an already-loaded entry within an explicit byte ceiling, checked before creating its defensive copy. External links are never resolved.</summary>
    public byte[] GetPackageEntryBytes(string path, long maximumBytes) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Package entry path cannot be empty.", nameof(path));
        if (maximumBytes < 0) throw new ArgumentOutOfRangeException(nameof(maximumBytes));
        string normalized = OdfPackagePath.NormalizeHref(path);
        if (normalized.Length == 0) throw new ArgumentException("Package entry path cannot be empty.", nameof(path));
        byte[] bytes = Package.GetRequiredEntry(normalized).GetBytesForSave();
        if (bytes.LongLength > maximumBytes) throw new InvalidDataException("Package entry exceeds maximumBytes.");
        return (byte[])bytes.Clone();
    }
}
