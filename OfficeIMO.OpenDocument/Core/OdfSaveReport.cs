namespace OfficeIMO.OpenDocument;

/// <summary>Describes which package entries a save rewrote, copied, removed, or could not project losslessly.</summary>
public sealed class OdfSaveReport {
    internal OdfSaveReport(IReadOnlyList<string> rewrittenEntries, IReadOnlyList<string> copiedEntries,
        IReadOnlyList<string> removedEntries, IReadOnlyList<string>? lossyEntries = null) {
        RewrittenEntries = rewrittenEntries;
        CopiedEntries = copiedEntries;
        RemovedEntries = removedEntries;
        LossyEntries = lossyEntries ?? Array.Empty<string>();
    }

    /// <summary>Entries serialized from changed state.</summary>
    public IReadOnlyList<string> RewrittenEntries { get; }
    /// <summary>Entries copied from their original payload.</summary>
    public IReadOnlyList<string> CopiedEntries { get; }
    /// <summary>Entries omitted from the output.</summary>
    public IReadOnlyList<string> RemovedEntries { get; }
    /// <summary>Source entries or parts that could not be represented losslessly by the selected output form.</summary>
    public IReadOnlyList<string> LossyEntries { get; }
}
