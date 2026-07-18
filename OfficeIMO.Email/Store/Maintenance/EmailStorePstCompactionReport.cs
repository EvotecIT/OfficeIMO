namespace OfficeIMO.Email.Store;

/// <summary>Verified PST rewrite-compaction outcome.</summary>
public sealed class EmailStorePstCompactionReport {
    internal EmailStorePstCompactionReport(EmailStorePstCompactionPlan plan,
        EmailStorePstConversionReport conversion) {
        Plan = plan;
        Conversion = conversion;
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
    public bool HasDataLoss => Conversion.HasDataLoss || !IsVerified;
}
