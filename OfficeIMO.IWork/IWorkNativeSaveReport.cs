namespace OfficeIMO.IWork;

/// <summary>Preservation facts for one completed encoding or publication of a captured native source.</summary>
public sealed class IWorkNativeSaveReport {
    internal IWorkNativeSaveReport(IWorkDocumentKind kind, string sourceHash, string outputHash, int entryCount,
        int recordCount, IEnumerable<IWorkObjectIdentity> updatedRecords, IEnumerable<string> updatedEntries,
        IEnumerable<IWorkDiagnostic> diagnostics) {
        Kind = kind; SourcePackageContentHash = sourceHash; OutputPackageContentHash = outputHash;
        EntryCount = entryCount; RecordCount = recordCount;
        UpdatedRecords = Array.AsReadOnly(updatedRecords.ToArray());
        UpdatedEntryPaths = Array.AsReadOnly(updatedEntries.OrderBy(path => path, StringComparer.Ordinal).ToArray());
        Diagnostics = Array.AsReadOnly(diagnostics.ToArray());
    }

    /// <summary>Gets the native application kind.</summary>
    public IWorkDocumentKind Kind { get; }
    /// <summary>Gets SHA-256 of the captured normalized source entries, including nested index expansion.</summary>
    public string SourcePackageContentHash { get; }
    /// <summary>Gets SHA-256 of the independently reopened output's normalized entries.</summary>
    public string OutputPackageContentHash { get; }
    /// <summary>Gets the physical file-entry count written to the ZIP carrier, excluding nested expansion.</summary>
    public int EntryCount { get; }
    /// <summary>Gets the preserved IWA record count, including auxiliary payloads.</summary>
    public int RecordCount { get; }
    /// <summary>Gets records whose image-position fields changed; all other payload bytes and headers remain intact.</summary>
    public IReadOnlyList<IWorkObjectIdentity> UpdatedRecords { get; }
    /// <summary>Gets compressed IWA entries rewritten by supported edits; other captured entry payloads remain byte-identical.</summary>
    public IReadOnlyList<string> UpdatedEntryPaths { get; }
    /// <summary>Gets carrier and native visual-cache limitations. This is separate from semantic conversion fidelity.</summary>
    public IReadOnlyList<IWorkDiagnostic> Diagnostics { get; }
}

/// <summary>A completed native ZIP encoding with its operation-level preservation report.</summary>
public sealed class IWorkNativeSaveResult {
    internal IWorkNativeSaveResult(byte[] bytes, IWorkNativeSaveReport report) { Bytes = bytes; Report = report; }
    internal byte[] Bytes { get; }
    /// <summary>Gets the encoded ZIP length.</summary>
    public int Length => Bytes.Length;
    /// <summary>Gets preservation and cache facts for this encoding.</summary>
    public IWorkNativeSaveReport Report { get; }
    /// <summary>Returns a defensive copy of the encoded native package.</summary>
    public byte[] GetPackageBytes() => (byte[])Bytes.Clone();
}
