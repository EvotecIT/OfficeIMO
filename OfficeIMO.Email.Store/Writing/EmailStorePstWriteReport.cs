namespace OfficeIMO.Email.Store;

/// <summary>Result of creating a new Unicode PST.</summary>
public sealed class EmailStorePstWriteReport {
    internal EmailStorePstWriteReport(string destinationPath, int folderCount, int itemCount,
        long bytesWritten, IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        DestinationPath = destinationPath;
        FolderCount = folderCount;
        ItemCount = itemCount;
        BytesWritten = bytesWritten;
        Diagnostics = diagnostics;
    }

    /// <summary>Committed destination path.</summary>
    public string DestinationPath { get; }

    /// <summary>Number of user-visible folders written, excluding mandatory PST system folders.</summary>
    public int FolderCount { get; }

    /// <summary>Number of top-level items written.</summary>
    public int ItemCount { get; }

    /// <summary>Committed file length.</summary>
    public long BytesWritten { get; }

    /// <summary>Preservation and compatibility diagnostics.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }

    /// <summary>True when at least one error diagnostic was emitted.</summary>
    public bool HasErrors => Diagnostics.Any(item => item.Severity == EmailStoreDiagnosticSeverity.Error);

    /// <summary>True when at least one fidelity warning or error was emitted.</summary>
    public bool HasDataLoss => Diagnostics.Any(item =>
        item.Severity == EmailStoreDiagnosticSeverity.Warning ||
        item.Severity == EmailStoreDiagnosticSeverity.Error);
}
