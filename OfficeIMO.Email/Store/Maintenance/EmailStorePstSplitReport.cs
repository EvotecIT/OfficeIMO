namespace OfficeIMO.Email.Store;

/// <summary>Verified outcome for one committed PST split part.</summary>
public sealed class EmailStorePstSplitPartReport {
    internal EmailStorePstSplitPartReport(EmailStorePstSplitPlanPart plan,
        EmailStorePstWriteReport writeReport, EmailStorePstVerificationReport verification,
        int skippedItems, IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        Plan = plan;
        WriteReport = writeReport;
        Verification = verification;
        SkippedItems = skippedItems;
        Diagnostics = diagnostics;
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
}

/// <summary>Aggregate verified query/size-based PST split result.</summary>
public sealed class EmailStorePstSplitReport {
    internal EmailStorePstSplitReport(EmailStorePstSplitPlan plan,
        IReadOnlyList<EmailStorePstSplitPartReport> parts,
        IReadOnlyList<EmailStoreDiagnostic> diagnostics) {
        Plan = plan;
        Parts = parts;
        Diagnostics = diagnostics;
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
}
