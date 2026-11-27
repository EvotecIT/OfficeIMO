namespace OfficeIMO.Email.Store;

/// <summary>Per-item result of recovering a damaged DBX or OLM source into a directory.</summary>
public sealed class EmailStoreSalvageReport : OfficeConversionReport {
    private readonly IReadOnlyList<OfficeDiagnostic> _fidelityDiagnostics;

    internal EmailStoreSalvageReport(EmailStoreFormat sourceFormat, string destinationDirectory,
        string? manifestPath, bool stoppedAtLimit, IReadOnlyList<EmailStoreSalvageEntry> entries,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        SourceFormat = sourceFormat;
        DestinationDirectory = destinationDirectory;
        ManifestPath = manifestPath;
        StoppedAtLimit = stoppedAtLimit;
        Entries = entries;
        NativeDiagnostics = diagnostics;
        var fidelity = new List<OfficeDiagnostic>(EmailStoreFidelityProjection.Project(diagnostics));
        foreach (EmailStoreSalvageEntry entry in entries) {
            if (entry.Status == EmailStoreSalvageStatus.Recovered) continue;
            fidelity.Add(EmailStoreFidelityProjection.Create(
                entry.Status == EmailStoreSalvageStatus.Partial
                    ? "EMAIL_STORE_SALVAGE_ITEM_PARTIAL"
                    : "EMAIL_STORE_SALVAGE_ITEM_NOT_RECOVERED",
                entry.Status == EmailStoreSalvageStatus.Partial
                    ? "Only part of the source item could be recovered."
                    : "The source item could not be recovered.",
                entry.Status == EmailStoreSalvageStatus.Partial
                    ? OfficeConversionLossKind.Approximation
                    : OfficeConversionLossKind.Failure,
                entry.ItemId));
        }
        if (stoppedAtLimit) {
            fidelity.Add(EmailStoreFidelityProjection.Create("EMAIL_STORE_SALVAGE_TRUNCATED",
                "The configured item bound stopped recovery before every candidate item was examined.",
                OfficeConversionLossKind.Omission, destinationDirectory));
        }
        _fidelityDiagnostics = fidelity.AsReadOnly();
    }

    /// <summary>Format of the damaged source.</summary>
    public EmailStoreFormat SourceFormat { get; }
    /// <summary>Absolute recovery root.</summary>
    public string DestinationDirectory { get; }
    /// <summary>Committed provenance manifest path, or null when disabled or unsuccessful.</summary>
    public string? ManifestPath { get; }
    /// <summary>Whether the item bound stopped recovery before every candidate was examined.</summary>
    public bool StoppedAtLimit { get; }
    /// <summary>Per-item outcomes in source order.</summary>
    public IReadOnlyList<EmailStoreSalvageEntry> Entries { get; }
    /// <summary>Source-level damage, limit, and publication diagnostics.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> NativeDiagnostics { get; }
    /// <summary>Items recovered completely.</summary>
    public int RecoveredCount => Entries.Count(static entry => entry.Status == EmailStoreSalvageStatus.Recovered);
    /// <summary>Items recovered with damaged, missing, or unverified parts.</summary>
    public int PartialCount => Entries.Count(static entry => entry.Status == EmailStoreSalvageStatus.Partial);
    /// <summary>Items found in the source that produced no artifact.</summary>
    public int NotRecoveredCount => Entries.Count(static entry => entry.Status == EmailStoreSalvageStatus.NotRecovered);
    /// <summary>Whether every candidate was examined and recovered completely.</summary>
    public bool IsComplete => !StoppedAtLimit && Entries.All(static entry => entry.Status == EmailStoreSalvageStatus.Recovered);
    /// <summary>Whether any item failed or an error diagnostic was reported.</summary>
    public bool HasErrors => NotRecoveredCount > 0 || NativeDiagnostics.Any(static diagnostic =>
        diagnostic.Severity == EmailStoreDiagnosticSeverity.Error);

    /// <summary>Source damage, partial items, failed items, and truncation as fidelity diagnostics.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => _fidelityDiagnostics;

    /// <summary>Whether recovery omitted, approximated, or failed to preserve source content.</summary>
    public override bool HasLoss => _fidelityDiagnostics.Any(static diagnostic =>
        diagnostic.LossKind != OfficeConversionLossKind.None);
}
