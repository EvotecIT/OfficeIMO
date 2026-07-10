namespace OfficeIMO.OpenDocument;

public abstract partial class OdfDocument {
    /// <summary>Package entry paths currently stored in the document.</summary>
    public IReadOnlyList<string> PackageEntries => Package.Entries.Select(entry => entry.Name).ToList();

    /// <summary>Returns a defensive copy of one already-loaded package entry without resolving external links.</summary>
    public byte[] GetPackageEntryBytes(string path) {
        if (string.IsNullOrWhiteSpace(path)) throw new ArgumentException("Package entry path cannot be empty.", nameof(path));
        byte[] bytes = Package.GetRequiredEntry(path).GetBytesForSave();
        return (byte[])bytes.Clone();
    }
}
