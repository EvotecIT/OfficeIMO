namespace OfficeIMO.Email.Store;

/// <summary>Non-mutating discovery of indexed items absent from normal folder contents tables.</summary>
public sealed class EmailStoreRecoveryReport {
    internal EmailStoreRecoveryReport(int itemsScanned, bool stoppedAtLimit,
        IReadOnlyList<EmailStoreItemReference> recoveredItems,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        ItemsScanned = itemsScanned;
        StoppedAtLimit = stoppedAtLimit;
        RecoveredItems = recoveredItems;
        Diagnostics = diagnostics;
    }

    /// <summary>Number of normal and recovered references examined.</summary>
    public int ItemsScanned { get; }

    /// <summary>Whether discovery stopped at a scan or result bound.</summary>
    public bool StoppedAtLimit { get; }

    /// <summary>Recovered stable references. Reading them remains an explicit operation.</summary>
    public IReadOnlyList<EmailStoreItemReference> RecoveredItems { get; }

    /// <summary>Diagnostics observed during index traversal.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }
}
