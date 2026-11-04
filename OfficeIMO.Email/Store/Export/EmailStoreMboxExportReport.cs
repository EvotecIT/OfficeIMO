namespace OfficeIMO.Email.Store;

/// <summary>Aggregate result of a streaming, atomically committed mbox export.</summary>
public sealed class EmailStoreMboxExportReport : OfficeConversionReport {
    private readonly IReadOnlyList<OfficeDiagnostic> _fidelityDiagnostics;

    internal EmailStoreMboxExportReport(string? destinationPath, bool wasTruncated,
        IReadOnlyList<EmailStoreMboxExportEntry> entries,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        DestinationPath = destinationPath;
        WasTruncated = wasTruncated;
        Entries = entries;
        NativeDiagnostics = diagnostics;
        int failedCount = entries.Count(static entry => !entry.Succeeded);
        _fidelityDiagnostics = EmailStoreFidelityProjection.ProjectExport(
            diagnostics,
            entries.Select(static entry => entry.Diagnostics),
            wasTruncated,
            "EMAIL_STORE_MBOX_SELECTION_TRUNCATED",
            "The configured item bound stopped mbox export before every selected source item was attempted.",
            failedCount,
            "EMAIL_STORE_MBOX_ITEMS_OMITTED",
            failedCount + " selected source item(s) were not appended to the mbox artifact.",
            destinationPath);
    }

    /// <summary>Absolute committed mbox path, or null when commit did not occur.</summary>
    public string? DestinationPath { get; }

    /// <summary>Whether export stopped at the configured item bound.</summary>
    public bool WasTruncated { get; }

    /// <summary>Per-item append outcomes.</summary>
    public IReadOnlyList<EmailStoreMboxExportEntry> Entries { get; }

    /// <summary>Session and destination diagnostics.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> NativeDiagnostics { get; }

    /// <summary>Successfully appended message count.</summary>
    public int SucceededCount => Entries.Count(item => item.Succeeded);

    /// <summary>Failed source item count.</summary>
    public int FailedCount => Entries.Count - SucceededCount;

    /// <summary>Total bytes appended to the mbox stream.</summary>
    public long BytesWritten => Entries.Sum(item => item.BytesWritten);

    /// <summary>Whether any destination, session, or item error occurred.</summary>
    public bool HasErrors => DestinationPath == null ||
        NativeDiagnostics.Any(item => item.Severity == EmailStoreDiagnosticSeverity.Error) ||
        Entries.Any(item => !item.Succeeded);

    /// <summary>Category-preserving session, item, truncation, and publication diagnostics.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics => _fidelityDiagnostics;

    /// <summary>Whether the export approximated, omitted, or failed to preserve selected source content.</summary>
    public override bool HasLoss => _fidelityDiagnostics.Any(static diagnostic =>
        diagnostic.LossKind != OfficeConversionLossKind.None);

}
