namespace OfficeIMO.OpenDocument;

/// <summary>Describes which package entries a save rewrote, copied, or removed.</summary>
public sealed class OdfSaveReport {
    internal OdfSaveReport(IReadOnlyList<string> rewrittenEntries, IReadOnlyList<string> copiedEntries, IReadOnlyList<string> removedEntries) {
        RewrittenEntries = rewrittenEntries;
        CopiedEntries = copiedEntries;
        RemovedEntries = removedEntries;
    }

    /// <summary>Entries serialized from changed state.</summary>
    public IReadOnlyList<string> RewrittenEntries { get; }
    /// <summary>Entries copied from their original payload.</summary>
    public IReadOnlyList<string> CopiedEntries { get; }
    /// <summary>Entries omitted from the output.</summary>
    public IReadOnlyList<string> RemovedEntries { get; }
}
