namespace OfficeIMO.Email.Store;

/// <summary>Aggregate result of an item-by-item store export.</summary>
public sealed class EmailStoreExportReport {
    internal EmailStoreExportReport(string destinationDirectory, bool wasTruncated,
        string? manifestPath, IReadOnlyList<EmailStoreExportEntry> entries,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        DestinationDirectory = destinationDirectory;
        WasTruncated = wasTruncated;
        ManifestPath = manifestPath;
        Entries = entries;
        Diagnostics = diagnostics;
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
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }

    /// <summary>Number of successfully written artifacts.</summary>
    public int SucceededCount => Entries.Count(item => item.Succeeded);

    /// <summary>Number of items that did not produce a valid artifact.</summary>
    public int FailedCount => Entries.Count - SucceededCount;

    /// <summary>Total serialized artifact bytes.</summary>
    public long BytesWritten => Entries.Sum(item => item.BytesWritten);

    /// <summary>Whether any session, manifest, or item error was reported.</summary>
    public bool HasErrors => Diagnostics.Any(item => item.Severity == EmailStoreDiagnosticSeverity.Error) ||
        Entries.Any(item => !item.Succeeded);
}
