namespace OfficeIMO.Email.Store;

/// <summary>Corruption-tolerant export result for indexed items absent from normal folder tables.</summary>
public sealed class EmailStoreRecoveryExportReport {
    internal EmailStoreRecoveryExportReport(string destinationDirectory,
        EmailStoreRecoveryReport discovery, string? manifestPath,
        IReadOnlyList<EmailStoreExportEntry> entries,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        DestinationDirectory = destinationDirectory;
        Discovery = discovery;
        ManifestPath = manifestPath;
        Entries = entries;
        Diagnostics = diagnostics;
    }

    /// <summary>Absolute recovery-export root.</summary>
    public string DestinationDirectory { get; }
    /// <summary>Bounded non-mutating recovery discovery that supplied the item references.</summary>
    public EmailStoreRecoveryReport Discovery { get; }
    /// <summary>Committed recovery manifest path, or null when disabled or unsuccessful.</summary>
    public string? ManifestPath { get; }
    /// <summary>Per-recovered-item export outcomes.</summary>
    public IReadOnlyList<EmailStoreExportEntry> Entries { get; }
    /// <summary>Discovery, source, export, and manifest diagnostics.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }
    /// <summary>Successfully preserved recovered items.</summary>
    public int RecoveredCount => Entries.Count(entry => entry.Succeeded);
    /// <summary>Recovered references that could not be read or serialized.</summary>
    public int FailedCount => Entries.Count - RecoveredCount;
    /// <summary>Whether discovery covered its selected scope and every discovered item was preserved.</summary>
    public bool IsComplete => !Discovery.StoppedAtLimit && FailedCount == 0;
    /// <summary>Whether discovery, export, or manifest processing produced an error.</summary>
    public bool HasErrors => FailedCount > 0 || Diagnostics.Any(diagnostic =>
        diagnostic.Severity == EmailStoreDiagnosticSeverity.Error);
}
