namespace OfficeIMO.OpenDocument;

/// <summary>Describes which package entries a save rewrote, copied, removed, or could not project losslessly.</summary>
public sealed class OdfSaveReport : global::OfficeIMO.IOfficeConversionReport {
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

    /// <summary>Category-preserving entry diagnostics for composed save routes.</summary>
    public IReadOnlyList<global::OfficeIMO.OfficeConversionFidelityDiagnostic> FidelityDiagnostics =>
        Array.AsReadOnly(RemovedEntries.Select(entry => new global::OfficeIMO.OfficeConversionFidelityDiagnostic(
                "ODF_ENTRY_REMOVED", "The source package entry was removed.",
                global::OfficeIMO.OfficeConversionLossKind.Omission, "OfficeIMO.OpenDocument", entry))
            .Concat(LossyEntries.Select(entry => new global::OfficeIMO.OfficeConversionFidelityDiagnostic(
                "ODF_ENTRY_PROJECTED", "The source package entry was projected with reduced fidelity.",
                global::OfficeIMO.OfficeConversionLossKind.Approximation, "OfficeIMO.OpenDocument", entry)))
            .ToArray());

    /// <summary>True when an entry was removed or projected lossily.</summary>
    public bool HasLoss => RemovedEntries.Count != 0 || LossyEntries.Count != 0;

    /// <summary>Throws when an entry was removed or projected lossily.</summary>
    public void RequireNoLoss() {
        if (HasLoss) throw new InvalidOperationException("OpenDocument save reported entry-level fidelity loss.");
    }
}
