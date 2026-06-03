namespace OfficeIMO.Pdf;

/// <summary>
/// Wrapper-friendly readiness report for a requested generated-PDF compliance profile.
/// </summary>
public sealed class PdfComplianceReadinessReport {
    internal PdfComplianceReadinessReport(PdfComplianceProfile profile, string displayName, IReadOnlyList<PdfComplianceRequirement> requirements) {
        Guard.ComplianceProfile(profile, nameof(profile));
        Guard.NotNullOrWhiteSpace(displayName, nameof(displayName));
        Guard.NotNull(requirements, nameof(requirements));

        Profile = profile;
        DisplayName = displayName;
        Requirements = requirements;
    }

    /// <summary>Requested compliance profile.</summary>
    public PdfComplianceProfile Profile { get; }

    /// <summary>Human-readable compliance profile name.</summary>
    public string DisplayName { get; }

    /// <summary>True when every known requirement is satisfied.</summary>
    public bool IsReady {
        get {
            for (int i = 0; i < Requirements.Count; i++) {
                if (Requirements[i].Status != PdfComplianceRequirementStatus.Satisfied) {
                    return false;
                }
            }

            return true;
        }
    }

    /// <summary>All evaluated requirements.</summary>
    public IReadOnlyList<PdfComplianceRequirement> Requirements { get; }

    /// <summary>Requirements that are not satisfied by the supplied options.</summary>
    public IReadOnlyList<PdfComplianceRequirement> MissingRequirements => GetByStatus(PdfComplianceRequirementStatus.Missing);

    /// <summary>Requirements that need additional OfficeIMO.Pdf engine support.</summary>
    public IReadOnlyList<PdfComplianceRequirement> UnsupportedRequirements => GetByStatus(PdfComplianceRequirementStatus.Unsupported);

    /// <summary>Finds a requirement by stable identifier, or returns null when it was not part of the requested profile.</summary>
    public PdfComplianceRequirement? FindRequirement(string id) {
        Guard.NotNullOrWhiteSpace(id, nameof(id));
        for (int i = 0; i < Requirements.Count; i++) {
            if (string.Equals(Requirements[i].Id, id, StringComparison.Ordinal)) {
                return Requirements[i];
            }
        }

        return null;
    }

    private System.Collections.ObjectModel.ReadOnlyCollection<PdfComplianceRequirement> GetByStatus(PdfComplianceRequirementStatus status) {
        var matches = new List<PdfComplianceRequirement>();
        for (int i = 0; i < Requirements.Count; i++) {
            if (Requirements[i].Status == status) {
                matches.Add(Requirements[i]);
            }
        }

        return matches.AsReadOnly();
    }
}
