namespace OfficeIMO.Email.Store;

/// <summary>Outcome of converting a supported store into a new Unicode PST.</summary>
public sealed class EmailStorePstConversionReport {
    internal EmailStorePstConversionReport(EmailStoreFormat sourceFormat,
        EmailStorePstWriteReport writeReport, int sourceFolders, int convertedItems,
        int skippedItems, EmailStorePstVerificationReport? verification,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics, EmailStoreSourceIdentity sourceIdentity,
        bool wasResumed) {
        SourceFormat = sourceFormat;
        WriteReport = writeReport;
        SourceFolders = sourceFolders;
        ConvertedItems = convertedItems;
        SkippedItems = skippedItems;
        Verification = verification;
        Diagnostics = diagnostics;
        SourceIdentity = sourceIdentity;
        WasResumed = wasResumed;
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
    /// <summary>Post-write semantic verification, or null when verification was disabled.</summary>
    public EmailStorePstVerificationReport? Verification { get; }
    /// <summary>Combined conversion and PST writer diagnostics.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }
    /// <summary>Privacy-safe exact source identity checked before and after migration.</summary>
    public EmailStoreSourceIdentity SourceIdentity { get; }
    /// <summary>True when this run continued an integrity-checked migration checkpoint.</summary>
    public bool WasResumed { get; }
    /// <summary>Strict final loss disposition.</summary>
    public EmailStoreMigrationDisposition Disposition => SkippedItems == 0 && !HasDataLoss
        ? EmailStoreMigrationDisposition.Completed
        : EmailStoreMigrationDisposition.CompletedWithAcceptedLoss;
    /// <summary>True when the conversion emitted a warning or error.</summary>
    public bool HasDataLoss => Verification?.IsSuccessful == false || Diagnostics.Any(item =>
        item.Severity != EmailStoreDiagnosticSeverity.Information);
}
