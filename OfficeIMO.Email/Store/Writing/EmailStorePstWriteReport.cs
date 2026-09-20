namespace OfficeIMO.Email.Store;

/// <summary>Result of creating a new Unicode PST.</summary>
public sealed class EmailStorePstWriteReport : IOfficeConversionReport {
    private readonly IReadOnlyList<OfficeConversionFidelityDiagnostic> _fidelityDiagnostics;

    internal EmailStorePstWriteReport(string destinationPath, int folderCount, int itemCount,
        long bytesWritten, IReadOnlyList<EmailStoreDiagnostic> diagnostics,
        bool diagnosticsTruncated = false) {
        DestinationPath = destinationPath;
        FolderCount = folderCount;
        ItemCount = itemCount;
        BytesWritten = bytesWritten;
        Diagnostics = diagnostics;
        DiagnosticsTruncated = diagnosticsTruncated;
        _fidelityDiagnostics = EmailStoreFidelityProjection.Project(diagnostics);
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

    /// <summary>True when additional diagnostics existed beyond the configured retention limit.</summary>
    public bool DiagnosticsTruncated { get; }

    /// <summary>True when at least one error diagnostic was emitted.</summary>
    public bool HasErrors => Diagnostics.Any(item => item.Severity == EmailStoreDiagnosticSeverity.Error);

    /// <summary>Category-preserving Store write diagnostics.</summary>
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics => _fidelityDiagnostics;

    /// <summary>True when the write approximated, omitted, or failed to preserve source content.</summary>
    public bool HasLoss => _fidelityDiagnostics.Any(static diagnostic =>
        diagnostic.LossKind != OfficeConversionLossKind.None);

    /// <summary>True when at least one fidelity warning or error was emitted.</summary>
    public bool HasDataLoss => HasLoss;

    /// <summary>Throws when the write reported possible content loss.</summary>
    public void RequireNoLoss() => EmailStoreFidelityProjection.RequireNoLoss(_fidelityDiagnostics);
}
