namespace OfficeIMO.Email.Store;

/// <summary>Result of creating a new Unicode PST.</summary>
public sealed class EmailStorePstWriteReport : OfficeConversionReport {
    private readonly IReadOnlyList<OfficeDiagnostic> _fidelityDiagnostics;

    internal EmailStorePstWriteReport(string destinationPath, int folderCount, int itemCount,
        long bytesWritten, IReadOnlyList<EmailStoreDiagnostic> diagnostics,
        bool diagnosticsTruncated = false) {
        DestinationPath = destinationPath;
        FolderCount = folderCount;
        ItemCount = itemCount;
        BytesWritten = bytesWritten;
        NativeDiagnostics = diagnostics;
        DiagnosticsTruncated = diagnosticsTruncated;
        _fidelityDiagnostics = EmailStoreFidelityProjection.ProjectWrite(
            diagnostics, diagnosticsTruncated, destinationPath);
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
    public IReadOnlyList<EmailStoreDiagnostic> NativeDiagnostics { get; }

    /// <summary>True when additional diagnostics existed beyond the configured retention limit.</summary>
    public bool DiagnosticsTruncated { get; }

    /// <summary>True when at least one error diagnostic was emitted.</summary>
    public bool HasErrors => NativeDiagnostics.Any(item => item.Severity == EmailStoreDiagnosticSeverity.Error);

    /// <summary>Category-preserving Store write diagnostics.</summary>
    protected override IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> GetDiagnostics() => FidelityDiagnostics;
    private IReadOnlyList<global::OfficeIMO.OfficeDiagnostic> FidelityDiagnostics => _fidelityDiagnostics;

    /// <summary>True when the write approximated, omitted, or failed to preserve source content.</summary>
    public override bool HasLoss => _fidelityDiagnostics.Any(static diagnostic =>
        diagnostic.LossKind != OfficeConversionLossKind.None);

    /// <summary>True when at least one fidelity warning or error was emitted.</summary>
    public bool HasDataLoss => HasLoss;

}
