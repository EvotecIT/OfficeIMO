namespace OfficeIMO.Email.Store;

/// <summary>Outcome of converting a supported store into a new Unicode PST.</summary>
public sealed class EmailStorePstConversionReport {
    internal EmailStorePstConversionReport(EmailStoreFormat sourceFormat,
        EmailStorePstWriteReport writeReport, int sourceFolders, int convertedItems,
        int skippedItems, IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        SourceFormat = sourceFormat;
        WriteReport = writeReport;
        SourceFolders = sourceFolders;
        ConvertedItems = convertedItems;
        SkippedItems = skippedItems;
        Diagnostics = diagnostics;
    }

    /// <summary>Detected source format.</summary>
    public EmailStoreFormat SourceFormat { get; }
    /// <summary>Final PST creation report.</summary>
    public EmailStorePstWriteReport WriteReport { get; }
    /// <summary>Number of source folders considered.</summary>
    public int SourceFolders { get; }
    /// <summary>Number of items written.</summary>
    public int ConvertedItems { get; }
    /// <summary>Number of items skipped after a reported read or fidelity failure.</summary>
    public int SkippedItems { get; }
    /// <summary>Combined conversion and PST writer diagnostics.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }
    /// <summary>True when the conversion emitted a warning or error.</summary>
    public bool HasDataLoss => Diagnostics.Any(item =>
        item.Severity != EmailStoreDiagnosticSeverity.Information);
}
