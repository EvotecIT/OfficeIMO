namespace OfficeIMO.Email.Store;

/// <summary>Outcome of converting a supported store into a new Unicode PST.</summary>
public sealed class EmailStorePstConversionReport : IOfficeConversionReport {
    private readonly IReadOnlyList<OfficeConversionFidelityDiagnostic> _fidelityDiagnostics;

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
        var additional = new List<OfficeConversionFidelityDiagnostic>();
        if (skippedItems > 0) {
            additional.Add(EmailStoreFidelityProjection.Create(
                "EMAIL_STORE_ITEMS_SKIPPED",
                skippedItems + " source item(s) were skipped during conversion.",
                OfficeConversionLossKind.Omission,
                "source-items"));
        }
        if (verification?.IsSuccessful == false) {
            additional.Add(EmailStoreFidelityProjection.Create(
                "EMAIL_STORE_PST_VERIFICATION_FAILED",
                "Post-write PST verification did not preserve the selected source semantics.",
                OfficeConversionLossKind.Failure,
                "verification"));
        }
        IReadOnlyList<OfficeConversionFidelityDiagnostic> projected =
            EmailStoreFidelityProjection.AppendMissing(
                EmailStoreFidelityProjection.Project(diagnostics),
                writeReport.FidelityDiagnostics);
        _fidelityDiagnostics = EmailStoreFidelityProjection.Append(projected, additional.ToArray());
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
    public bool HasDataLoss => HasLoss;
    /// <summary>Category-preserving Store conversion diagnostics.</summary>
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics => _fidelityDiagnostics;
    /// <summary>True when conversion approximated, omitted, or failed to preserve source content.</summary>
    public bool HasLoss => _fidelityDiagnostics.Any(static diagnostic =>
        diagnostic.LossKind != OfficeConversionLossKind.None);
    /// <summary>Throws when conversion reported possible content loss.</summary>
    public void RequireNoLoss() => EmailStoreFidelityProjection.RequireNoLoss(_fidelityDiagnostics);
}
