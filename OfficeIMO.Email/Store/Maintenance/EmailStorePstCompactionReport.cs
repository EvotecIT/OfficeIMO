namespace OfficeIMO.Email.Store;

/// <summary>Verified PST rewrite-compaction outcome.</summary>
public sealed class EmailStorePstCompactionReport : IOfficeConversionReport {
    private readonly IReadOnlyList<OfficeConversionFidelityDiagnostic> _fidelityDiagnostics;

    internal EmailStorePstCompactionReport(EmailStorePstCompactionPlan plan,
        EmailStorePstConversionReport conversion) {
        Plan = plan;
        Conversion = conversion;
        _fidelityDiagnostics = Conversion.Verification?.IsSuccessful == true &&
            Conversion.ConvertedItems == Plan.SelectedItems && Conversion.SkippedItems == 0
            ? Conversion.FidelityDiagnostics
            : EmailStoreFidelityProjection.Append(
                Conversion.FidelityDiagnostics,
                EmailStoreFidelityProjection.Create(
                    "EMAIL_STORE_PST_COMPACTION_VERIFICATION_FAILED",
                    "The compacted PST did not contain and verify every selected item.",
                    OfficeConversionLossKind.Failure,
                    "verification"));
    }

    /// <summary>Pre-write selection and capacity plan.</summary>
    public EmailStorePstCompactionPlan Plan { get; }
    /// <summary>Existing verified conversion/rewrite report.</summary>
    public EmailStorePstConversionReport Conversion { get; }
    /// <summary>Committed compacted PST length.</summary>
    public long CompactedBytes => Conversion.WriteReport.BytesWritten;
    /// <summary>Actual source bytes minus compacted bytes. Negative means the canonical rewrite grew.</summary>
    public long ReductionBytes => Plan.SourceBytes - CompactedBytes;
    /// <summary>Actual fractional size reduction; negative means growth.</summary>
    public double ReductionRatio => Plan.SourceBytes == 0 ? 0 :
        (double)ReductionBytes / Plan.SourceBytes;
    /// <summary>Whether every selected item was written, reopened, and semantically matched.</summary>
    public bool IsVerified => Conversion.Verification?.IsSuccessful == true &&
        Conversion.ConvertedItems == Plan.SelectedItems && Conversion.SkippedItems == 0;
    /// <summary>Whether the rewrite reported preservation loss.</summary>
    public bool HasDataLoss => HasLoss;
    /// <summary>Category-preserving Store compaction diagnostics.</summary>
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> FidelityDiagnostics => _fidelityDiagnostics;
    /// <summary>True when compaction approximated, omitted, or failed to preserve source content.</summary>
    public bool HasLoss => _fidelityDiagnostics.Any(static diagnostic =>
        diagnostic.LossKind != OfficeConversionLossKind.None);
    /// <summary>Throws when compaction reported possible content loss.</summary>
    public void RequireNoLoss() => EmailStoreFidelityProjection.RequireNoLoss(_fidelityDiagnostics);
}
