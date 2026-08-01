namespace OfficeIMO.Email.Store;

/// <summary>Source-preserving maintenance action proposed by inspection and bounded validation.</summary>
public enum EmailStoreMaintenanceAction {
    /// <summary>No maintenance is currently indicated.</summary>
    None = 0,
    /// <summary>Export recoverable items to a distinct destination.</summary>
    RecoveryExport = 1,
    /// <summary>Rewrite a PST to a distinct destination and verify it semantically.</summary>
    VerifiedRewrite = 2,
    /// <summary>Split a PST into bounded, separately verified destinations.</summary>
    Split = 3,
    /// <summary>The source requires an external or manual repair tool.</summary>
    ManualIntervention = 4
}

/// <summary>One explainable recommendation in a read-only maintenance plan.</summary>
public sealed class EmailStoreMaintenanceRecommendation {
    internal EmailStoreMaintenanceRecommendation(EmailStoreMaintenanceAction action, string reason,
        bool executableByOfficeImo, EmailDataLossRisk dataLossRisk) {
        Action = action; Reason = reason; ExecutableByOfficeImo = executableByOfficeImo;
        DataLossRisk = dataLossRisk;
    }
    /// <summary>Proposed action.</summary>
    public EmailStoreMaintenanceAction Action { get; }
    /// <summary>Operator-facing reason.</summary>
    public string Reason { get; }
    /// <summary>Whether OfficeIMO has a plan-and-verify execution path for this action.</summary>
    public bool ExecutableByOfficeImo { get; }
    /// <summary>Potential data-loss classification if the recommendation is followed.</summary>
    public EmailDataLossRisk DataLossRisk { get; }
}

/// <summary>Read-only maintenance decision bound to one complete source fingerprint.</summary>
public sealed class EmailStoreMaintenancePlan {
    internal EmailStoreMaintenancePlan(string sourceFingerprint, EmailStoreFormat format,
        EmailStoreValidationReport validation, EmailStoreRecoveryReport recovery,
        IReadOnlyList<EmailStoreMaintenanceRecommendation> recommendations) {
        SourceFingerprint = sourceFingerprint; Format = format; Validation = validation;
        Recovery = recovery; Recommendations = recommendations;
    }
    /// <summary>Complete-source SHA-256 identity captured while planning.</summary>
    public string SourceFingerprint { get; }
    /// <summary>Detected source format.</summary>
    public EmailStoreFormat Format { get; }
    /// <summary>Bounded validation evidence used by the planner.</summary>
    public EmailStoreValidationReport Validation { get; }
    /// <summary>Bounded orphan/recovery evidence used by the planner.</summary>
    public EmailStoreRecoveryReport Recovery { get; }
    /// <summary>Ordered explainable recommendations.</summary>
    public IReadOnlyList<EmailStoreMaintenanceRecommendation> Recommendations { get; }
    /// <summary>Maintenance never writes to the opened source.</summary>
    public bool PreservesSource => true;
    /// <summary>OfficeIMO rewrite actions require a semantic post-write verification report.</summary>
    public bool RequiresPostWriteVerification => Recommendations.Any(item =>
        item.Action == EmailStoreMaintenanceAction.VerifiedRewrite || item.Action == EmailStoreMaintenanceAction.Split);
}
