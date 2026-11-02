namespace OfficeIMO.OpenDocument;

/// <summary>Describes which package entries a save rewrote, copied, removed, or could not project losslessly.</summary>
public sealed class OdfSaveReport : OfficeConversionReport {
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
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics =>
        Array.AsReadOnly(RemovedEntries.Select(entry => new global::OfficeIMO.OfficeDiagnostic(
                "ODF_ENTRY_REMOVED",
                global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
                "The source package entry was removed.",
                entry,
                "OfficeIMO.OpenDocument",
                global::OfficeIMO.OfficeConversionLossKind.Omission))
            .Concat(LossyEntries.Select(entry => new global::OfficeIMO.OfficeDiagnostic(
                "ODF_ENTRY_OMITTED",
                global::OfficeIMO.OfficeDiagnosticSeverity.Warning,
                "The source package entry was not represented in the flat document.",
                entry,
                "OfficeIMO.OpenDocument",
                global::OfficeIMO.OfficeConversionLossKind.Omission)))
            .ToArray());

    /// <summary>True when an entry was removed or projected lossily.</summary>
    public override bool HasLoss => RemovedEntries.Count != 0 || LossyEntries.Count != 0;

    /// <summary>Throws when an entry was removed or projected lossily.</summary>
    protected override OfficeConversionException CreateLossException() => new OfficeConversionException("OpenDocument save reported entry-level fidelity loss.", this);
}
