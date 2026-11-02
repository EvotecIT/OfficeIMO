namespace OfficeIMO.Email.Store;

/// <summary>Aggregate result of an item-by-item store export.</summary>
public sealed class EmailStoreExportReport : OfficeConversionReport {
    private readonly IReadOnlyList<OfficeDiagnostic> _fidelityDiagnostics;

    internal EmailStoreExportReport(string destinationDirectory, bool wasTruncated,
        string? manifestPath, IReadOnlyList<EmailStoreExportEntry> entries,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        DestinationDirectory = destinationDirectory;
        WasTruncated = wasTruncated;
        ManifestPath = manifestPath;
        Entries = entries;
        NativeDiagnostics = diagnostics;
        int failedCount = entries.Count(static entry => !entry.Succeeded);
        _fidelityDiagnostics = EmailStoreFidelityProjection.ProjectExport(
            diagnostics,
            entries.Select(static entry => entry.Diagnostics),
            wasTruncated,
            "EMAIL_STORE_EXPORT_SELECTION_TRUNCATED",
            "The configured item bound stopped export before every selected source item was attempted.",
            failedCount,
            "EMAIL_STORE_EXPORT_ITEMS_OMITTED",
            failedCount + " selected source item(s) did not produce a destination artifact.",
            destinationDirectory);
    }

    /// <summary>Absolute export root.</summary>
    public string DestinationDirectory { get; }

    /// <summary>Whether export stopped at the configured item bound.</summary>
    public bool WasTruncated { get; }

    /// <summary>Absolute manifest path when one was written.</summary>
    public string? ManifestPath { get; }

    /// <summary>Per-item export outcomes.</summary>
    public IReadOnlyList<EmailStoreExportEntry> Entries { get; }

    /// <summary>Session-level and manifest diagnostics.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> NativeDiagnostics { get; }

    /// <summary>Number of successfully written artifacts.</summary>
    public int SucceededCount => Entries.Count(item => item.Succeeded);

    /// <summary>Number of items that did not produce a valid artifact.</summary>
    public int FailedCount => Entries.Count - SucceededCount;

    /// <summary>Total serialized artifact bytes.</summary>
    public long BytesWritten => Entries.Sum(item => item.BytesWritten);

    /// <summary>Whether any session, manifest, or item error was reported.</summary>
    public bool HasErrors => NativeDiagnostics.Any(item => item.Severity == EmailStoreDiagnosticSeverity.Error) ||
        Entries.Any(item => !item.Succeeded);

    /// <summary>Category-preserving session, item, truncation, and publication diagnostics.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics => _fidelityDiagnostics;

    /// <summary>Whether the export approximated, omitted, or failed to preserve selected source content.</summary>
    public override bool HasLoss => _fidelityDiagnostics.Any(static diagnostic =>
        diagnostic.LossKind != OfficeConversionLossKind.None);

}
