namespace OfficeIMO.Pdf;

/// <summary>
/// One generated-PDF compliance requirement and its current readiness state.
/// </summary>
public sealed class PdfComplianceRequirement {
    internal PdfComplianceRequirement(string id, string displayName, PdfComplianceRequirementStatus status, string diagnostic) {
        Guard.NotNullOrWhiteSpace(id, nameof(id));
        Guard.NotNullOrWhiteSpace(displayName, nameof(displayName));
        Guard.NotNullOrWhiteSpace(diagnostic, nameof(diagnostic));

        Id = id;
        DisplayName = displayName;
        Status = status;
        Diagnostic = diagnostic;
    }

    /// <summary>Stable requirement identifier for wrapper diagnostics.</summary>
    public string Id { get; }

    /// <summary>Human-readable requirement name.</summary>
    public string DisplayName { get; }

    /// <summary>Current readiness state.</summary>
    public PdfComplianceRequirementStatus Status { get; }

    /// <summary>Human-readable explanation of the readiness state.</summary>
    public string Diagnostic { get; }
}
