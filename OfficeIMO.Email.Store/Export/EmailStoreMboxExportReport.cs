namespace OfficeIMO.Email.Store;

/// <summary>Aggregate result of a streaming, atomically committed mbox export.</summary>
public sealed class EmailStoreMboxExportReport {
    internal EmailStoreMboxExportReport(string? destinationPath, bool wasTruncated,
        IReadOnlyList<EmailStoreMboxExportEntry> entries,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        DestinationPath = destinationPath;
        WasTruncated = wasTruncated;
        Entries = entries;
        Diagnostics = diagnostics;
    }

    /// <summary>Absolute committed mbox path, or null when commit did not occur.</summary>
    public string? DestinationPath { get; }

    /// <summary>Whether export stopped at the configured item bound.</summary>
    public bool WasTruncated { get; }

    /// <summary>Per-item append outcomes.</summary>
    public IReadOnlyList<EmailStoreMboxExportEntry> Entries { get; }

    /// <summary>Session and destination diagnostics.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }

    /// <summary>Successfully appended message count.</summary>
    public int SucceededCount => Entries.Count(item => item.Succeeded);

    /// <summary>Failed source item count.</summary>
    public int FailedCount => Entries.Count - SucceededCount;

    /// <summary>Total bytes appended to the mbox stream.</summary>
    public long BytesWritten => Entries.Sum(item => item.BytesWritten);

    /// <summary>Whether any destination, session, or item error occurred.</summary>
    public bool HasErrors => DestinationPath == null ||
        Diagnostics.Any(item => item.Severity == EmailStoreDiagnosticSeverity.Error) ||
        Entries.Any(item => !item.Succeeded);
}
