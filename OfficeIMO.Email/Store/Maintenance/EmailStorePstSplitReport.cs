namespace OfficeIMO.Email.Store;

/// <summary>Verified outcome for one committed PST split part.</summary>
public sealed class EmailStorePstSplitPartReport : IOfficeConversionReport {
    private readonly IReadOnlyList<OfficeConversionFidelityDiagnostic> _fidelityDiagnostics;

    internal EmailStorePstSplitPartReport(EmailStorePstSplitPlanPart plan,
        EmailStorePstWriteReport writeReport, EmailStorePstVerificationReport verification,
        int skippedItems, IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        Plan = plan;
        WriteReport = writeReport;
        Verification = verification;
        SkippedItems = skippedItems;
        Diagnostics = diagnostics;
        var aggregate = new List<OfficeConversionFidelityDiagnostic>();
        if (skippedItems > 0) {
            aggregate.Add(EmailStoreFidelityProjection.Create(
                "EMAIL_STORE_PST_SPLIT_ITEMS_OMITTED",
                skippedItems + " selected source item(s) were omitted from this split part.",
                OfficeConversionLossKind.Omission,
                plan.DestinationPath));
        }
        if (!verification.IsSuccessful) {
            aggregate.Add(EmailStoreFidelityProjection.Create(
                "EMAIL_STORE_PST_SPLIT_VERIFICATION_FAILED",
                "The committed split part did not verify every written item against the selected source semantics.",
                OfficeConversionLossKind.Failure,
                plan.DestinationPath));
        }
        _fidelityDiagnostics = EmailStoreFidelityProjection.Append(
            EmailStoreFidelityProjection.Project(diagnostics), aggregate.ToArray());
    }

    /// <summary>Dry-run partition that produced this part.</summary>
    public EmailStorePstSplitPlanPart Plan { get; }
    /// <summary>Committed PST writer report.</summary>
    public EmailStorePstWriteReport WriteReport { get; }
    /// <summary>Mandatory reopen-and-semantic-compare verification.</summary>
    public EmailStorePstVerificationReport Verification { get; }
    /// <summary>Selected items skipped under explicit continuation policy.</summary>
    public int SkippedItems { get; }
    /// <summary>Part-specific preservation and verification diagnostics.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }
    /// <summary>Final bytes minus the dry-run estimate. Positive means the final PST was larger.</summary>
    public long EstimateDeltaBytes => WriteReport.BytesWritten - Plan.EstimatedBytes;
    /// <summary>Whether the final PST exceeded the configured estimated partition target.</summary>
    public bool ExceededEstimatedTarget => WriteReport.BytesWritten > Plan.EstimatedTargetBytes;
    /// <summary>Whether every written item was reopened and matched.</summary>
    public bool IsVerified => Verification.IsSuccessful;

    /// <summary>Category-preserving part, omission, writer, and verification diagnostics.</summary>
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics => _fidelityDiagnostics;

    /// <summary>Whether this part approximated, omitted, or failed to preserve selected source content.</summary>
    public bool HasLoss => _fidelityDiagnostics.Any(static diagnostic =>
        diagnostic.LossKind != OfficeConversionLossKind.None);

    /// <summary>Throws when this split part reported possible content loss.</summary>
    public void RequireNoLoss() => EmailStoreFidelityProjection.RequireNoLoss(_fidelityDiagnostics);
}

/// <summary>Aggregate verified query/size-based PST split result.</summary>
public sealed class EmailStorePstSplitReport : IOfficeConversionReport {
    private readonly IReadOnlyList<OfficeConversionFidelityDiagnostic> _fidelityDiagnostics;

    internal EmailStorePstSplitReport(EmailStorePstSplitPlan plan,
        IReadOnlyList<EmailStorePstSplitPartReport> parts,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        Plan = plan;
        Parts = parts;
        Diagnostics = diagnostics;
        IReadOnlyList<OfficeConversionFidelityDiagnostic> projected =
            OfficeConversionFidelityDiagnostics.Flatten(parts);
        var aggregate = new List<OfficeConversionFidelityDiagnostic>();
        if (parts.Count != plan.Parts.Count) {
            aggregate.Add(EmailStoreFidelityProjection.Create(
                "EMAIL_STORE_PST_SPLIT_PARTS_MISSING",
                "The split operation did not commit every planned output part.",
                OfficeConversionLossKind.Failure,
                plan.OutputBasePath));
        }
        _fidelityDiagnostics = EmailStoreFidelityProjection.Append(projected, aggregate.ToArray());
    }

    /// <summary>Executed dry-run plan.</summary>
    public EmailStorePstSplitPlan Plan { get; }
    /// <summary>Committed verified parts.</summary>
    public IReadOnlyList<EmailStorePstSplitPartReport> Parts { get; }
    /// <summary>Aggregate commit and fidelity diagnostics.</summary>
    public IReadOnlyList<EmailStoreDiagnostic> Diagnostics { get; }
    /// <summary>Total committed items.</summary>
    public int WrittenItems => Parts.Sum(part => part.WriteReport.ItemCount);
    /// <summary>Total committed output bytes.</summary>
    public long BytesWritten => Parts.Sum(part => part.WriteReport.BytesWritten);
    /// <summary>True when every planned part was committed and semantically verified.</summary>
    public bool IsSuccessful => Parts.Count == Plan.Parts.Count &&
        Parts.All(part => part.IsVerified) &&
        !Diagnostics.Any(diagnostic => diagnostic.Severity == EmailStoreDiagnosticSeverity.Error);

    /// <summary>Category-preserving aggregate and per-part split diagnostics.</summary>
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics => _fidelityDiagnostics;

    /// <summary>Whether split approximated, omitted, or failed to preserve selected source content.</summary>
    public bool HasLoss => _fidelityDiagnostics.Any(static diagnostic =>
        diagnostic.LossKind != OfficeConversionLossKind.None);

    /// <summary>Throws when split reported possible content loss.</summary>
    public void RequireNoLoss() => EmailStoreFidelityProjection.RequireNoLoss(_fidelityDiagnostics);
}
