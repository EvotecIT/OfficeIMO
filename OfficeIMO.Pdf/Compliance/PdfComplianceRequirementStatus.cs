namespace OfficeIMO.Pdf;

/// <summary>
/// Readiness state for one generated-PDF compliance requirement.
/// </summary>
public enum PdfComplianceRequirementStatus {
    /// <summary>The requirement is satisfied by the supplied generation options.</summary>
    Satisfied,
    /// <summary>The requirement is not satisfied by the supplied generation options.</summary>
    Missing,
    /// <summary>The requirement needs engine support that is not implemented yet.</summary>
    Unsupported
}
