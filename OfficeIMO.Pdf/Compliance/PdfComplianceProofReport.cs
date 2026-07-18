namespace OfficeIMO.Pdf;

/// <summary>
/// Combines OfficeIMO.Pdf readiness diagnostics with caller-supplied external validator evidence.
/// </summary>
public sealed class PdfComplianceProofReport {
    private readonly IReadOnlyList<PdfExternalValidatorKind> _requiredExternalValidators;
    private readonly IReadOnlyList<PdfExternalValidationResult> _externalValidations;

    internal PdfComplianceProofReport(
        PdfComplianceReadinessReport readiness,
        IReadOnlyList<PdfExternalValidatorKind> requiredExternalValidators,
        IReadOnlyList<PdfExternalValidationResult> externalValidations,
        string? artifactSha256 = null,
        long? artifactSizeBytes = null) {
        Guard.NotNull(readiness, nameof(readiness));
        Guard.NotNull(requiredExternalValidators, nameof(requiredExternalValidators));
        Guard.NotNull(externalValidations, nameof(externalValidations));

        Readiness = readiness;
        _requiredExternalValidators = requiredExternalValidators;
        _externalValidations = externalValidations;
        ArtifactSha256 = PdfArtifactFingerprint.NormalizeSha256(artifactSha256, nameof(artifactSha256));
        ArtifactSizeBytes = artifactSizeBytes;
    }

    /// <summary>Requested compliance profile.</summary>
    public PdfComplianceProfile Profile => Readiness.Profile;

    /// <summary>Human-readable compliance profile name.</summary>
    public string DisplayName => Readiness.DisplayName;

    /// <summary>OfficeIMO.Pdf readiness report, including non-external requirements and external-evidence placeholders.</summary>
    public PdfComplianceReadinessReport Readiness { get; }

    /// <summary>External validator families required before the requested profile can be claimed.</summary>
    public IReadOnlyList<PdfExternalValidatorKind> RequiredExternalValidators => _requiredExternalValidators;

    /// <summary>Caller-supplied external validation results.</summary>
    public IReadOnlyList<PdfExternalValidationResult> ExternalValidations => _externalValidations;

    /// <summary>
    /// De-duplicated requirements after external validator evidence has been reconciled.
    /// Passing external proof rows are represented as satisfied requirements.
    /// </summary>
    public IReadOnlyList<PdfComplianceRequirement> EffectiveRequirements {
        get {
            var requirements = new List<PdfComplianceRequirement>(Readiness.Requirements.Count);
            var ids = new HashSet<string>(StringComparer.Ordinal);
            for (int i = 0; i < Readiness.Requirements.Count; i++) {
                PdfComplianceRequirement requirement = Readiness.Requirements[i];
                if (!ids.Add(requirement.Id)) {
                    continue;
                }

                if (TryGetExternalValidator(requirement.Id, out PdfExternalValidatorKind validatorKind) &&
                    FindExternalValidatorProof(validatorKind)?.IsSatisfied == true) {
                    requirements.Add(new PdfComplianceRequirement(
                        requirement.Id,
                        requirement.DisplayName,
                        PdfComplianceRequirementStatus.Satisfied,
                        "Required external validation passed for the exact artifact."));
                } else {
                    requirements.Add(requirement);
                }
            }

            return requirements.AsReadOnly();
        }
    }

    /// <summary>Effective requirements still missing after external evidence reconciliation.</summary>
    public IReadOnlyList<PdfComplianceRequirement> MissingRequirements =>
        EffectiveRequirements
            .Where(static requirement => requirement.Status == PdfComplianceRequirementStatus.Missing)
            .ToArray();

    /// <summary>Effective requirements still unsupported after external evidence reconciliation.</summary>
    public IReadOnlyList<PdfComplianceRequirement> UnsupportedRequirements =>
        EffectiveRequirements
            .Where(static requirement => requirement.Status == PdfComplianceRequirementStatus.Unsupported)
            .ToArray();

    /// <summary>All effective requirements that currently block a conformance claim.</summary>
    public IReadOnlyList<PdfComplianceRequirement> BlockingRequirements =>
        EffectiveRequirements
            .Where(static requirement => requirement.Status != PdfComplianceRequirementStatus.Satisfied)
            .ToArray();

    /// <summary>SHA-256 of the exact PDF artifact represented by this proof report.</summary>
    public string? ArtifactSha256 { get; }

    /// <summary>Size of the exact PDF artifact represented by this proof report, in bytes.</summary>
    public long? ArtifactSizeBytes { get; }

    /// <summary>True when the report represents one exact PDF artifact.</summary>
    public bool HasArtifactEvidence => ArtifactSha256 != null && ArtifactSizeBytes.HasValue;

    /// <summary>One machine-readable proof row for each external validator family required by the requested profile.</summary>
    public IReadOnlyList<PdfExternalValidatorProof> ExternalValidatorProofs {
        get {
            var proofs = new List<PdfExternalValidatorProof>();
            for (int i = 0; i < _requiredExternalValidators.Count; i++) {
                proofs.Add(BuildExternalValidatorProof(_requiredExternalValidators[i]));
            }

            return proofs.AsReadOnly();
        }
    }

    /// <summary>True when every non-external OfficeIMO.Pdf readiness requirement is satisfied.</summary>
    public bool IsInternallyReady {
        get {
            for (int i = 0; i < Readiness.Requirements.Count; i++) {
                PdfComplianceRequirement requirement = Readiness.Requirements[i];
                if (IsExternalValidationRequirement(requirement.Id)) {
                    continue;
                }

                if (requirement.Status != PdfComplianceRequirementStatus.Satisfied) {
                    return false;
                }
            }

            return true;
        }
    }

    /// <summary>True when every required external validator family has a passing result.</summary>
    public bool HasRequiredExternalValidation {
        get {
            for (int i = 0; i < _requiredExternalValidators.Count; i++) {
                PdfExternalValidatorKind validator = _requiredExternalValidators[i];
                if (!HasPassingExternalValidation(validator) || HasFailedExternalValidation(validator)) {
                    return false;
                }
            }

            return true;
        }
    }

    /// <summary>True only when internal readiness is satisfied, every required external validator passed, and no required validator failed.</summary>
    public bool CanClaimConformance =>
        Profile != PdfComplianceProfile.None &&
        HasArtifactEvidence &&
        IsInternallyReady &&
        HasRequiredExternalValidation &&
        FailedExternalValidations.Count == 0 &&
        BlockingRequirements.Count == 0;

    /// <summary>True when OfficeIMO.Pdf readiness is complete and the remaining proof work is external validation.</summary>
    public bool ReadyForExternalValidation =>
        Profile != PdfComplianceProfile.None &&
        HasArtifactEvidence &&
        IsInternallyReady &&
        RequiresExternalValidation &&
        !HasRequiredExternalValidation &&
        FailedExternalValidationCount == 0;

    /// <summary>Number of required external validator families.</summary>
    public int RequiredExternalValidatorCount => _requiredExternalValidators.Count;

    /// <summary>Number of caller-supplied passing validation results for the requested profile and required validator families.</summary>
    public int PassedExternalValidationCount => CountExternalValidations(PdfExternalValidationStatus.Passed);

    /// <summary>Number of caller-supplied failed or errored validation results for the requested profile and required validator families.</summary>
    public int FailedExternalValidationCount => FailedExternalValidations.Count;

    /// <summary>Number of required validator families without a current passing result.</summary>
    public int MissingExternalValidatorCount => MissingExternalValidators.Count;

    /// <summary>True when the profile requires external validation before conformance can be claimed.</summary>
    public bool RequiresExternalValidation => RequiredExternalValidatorCount > 0;

    /// <summary>Stable proof state for automation: None, InternalGaps, MissingArtifactEvidence, MissingExternalValidation, ExternalValidationFailed, or Claimable.</summary>
    public string ProofStatus {
        get {
            if (Profile == PdfComplianceProfile.None) {
                return "None";
            }

            if (!IsInternallyReady) {
                return "InternalGaps";
            }

            if (!HasArtifactEvidence) {
                return "MissingArtifactEvidence";
            }

            if (FailedExternalValidationCount > 0) {
                return "ExternalValidationFailed";
            }

            if (!HasRequiredExternalValidation) {
                return "MissingExternalValidation";
            }

            return "Claimable";
        }
    }

    /// <summary>Human-readable external proof summary suitable for command-line output.</summary>
    public string ExternalProofSummary {
        get {
            if (!RequiresExternalValidation) {
                return "No external validator is required for this profile.";
            }

            if (FailedExternalValidationCount > 0) {
                return "External validation failed or errored: " + string.Join(", ", FailedExternalValidations.Select(static result => result.ValidatorName));
            }

            if (HasRequiredExternalValidation) {
                return "Required external validators passed.";
            }

            return "Missing external validation: " + string.Join(", ", MissingExternalValidators.Select(static validator => validator.ToString()));
        }
    }

    /// <summary>Required external validators that do not have a passing result.</summary>
    public IReadOnlyList<PdfExternalValidatorKind> MissingExternalValidators {
        get {
            var missing = new List<PdfExternalValidatorKind>();
            for (int i = 0; i < _requiredExternalValidators.Count; i++) {
                PdfExternalValidatorKind validator = _requiredExternalValidators[i];
                if (!HasPassingExternalValidation(validator) || HasFailedExternalValidation(validator)) {
                    missing.Add(validator);
                }
            }

            return missing.AsReadOnly();
        }
    }

    /// <summary>Required external validator results that failed or errored.</summary>
    public IReadOnlyList<PdfExternalValidationResult> FailedExternalValidations {
        get {
            var failed = new List<PdfExternalValidationResult>();
            for (int i = 0; i < _externalValidations.Count; i++) {
                PdfExternalValidationResult result = _externalValidations[i];
                if (!_requiredExternalValidators.Contains(result.ValidatorKind) ||
                    !IsExternalValidationForRequestedProfile(result) ||
                    !IsExternalValidationForArtifact(result)) {
                    continue;
                }

                if (result.Status == PdfExternalValidationStatus.Failed ||
                    result.Status == PdfExternalValidationStatus.Error) {
                    failed.Add(result);
                }
            }

            return failed.AsReadOnly();
        }
    }

    /// <summary>Finds the first external validation result for the requested validator family.</summary>
    public PdfExternalValidationResult? FindExternalValidation(PdfExternalValidatorKind validatorKind) {
        for (int i = 0; i < _externalValidations.Count; i++) {
            PdfExternalValidationResult result = _externalValidations[i];
            if (result.ValidatorKind == validatorKind &&
                IsExternalValidationForRequestedProfile(result) &&
                IsExternalValidationForArtifact(result)) {
                return result;
            }
        }

        return null;
    }

    /// <summary>Finds the proof row for a required external validator family.</summary>
    public PdfExternalValidatorProof? FindExternalValidatorProof(PdfExternalValidatorKind validatorKind) {
        for (int i = 0; i < _requiredExternalValidators.Count; i++) {
            if (_requiredExternalValidators[i] == validatorKind) {
                return BuildExternalValidatorProof(validatorKind);
            }
        }

        return null;
    }

    internal static bool IsExternalValidationRequirement(string id) =>
        string.Equals(id, "verapdf-validation", StringComparison.Ordinal) ||
        string.Equals(id, "pdfua-validation", StringComparison.Ordinal) ||
        string.Equals(id, "mustang-validation", StringComparison.Ordinal);

    private static bool TryGetExternalValidator(string id, out PdfExternalValidatorKind validatorKind) {
        if (string.Equals(id, "verapdf-validation", StringComparison.Ordinal)) {
            validatorKind = PdfExternalValidatorKind.VeraPdf;
            return true;
        }

        if (string.Equals(id, "pdfua-validation", StringComparison.Ordinal)) {
            validatorKind = PdfExternalValidatorKind.PdfUaValidator;
            return true;
        }

        if (string.Equals(id, "mustang-validation", StringComparison.Ordinal)) {
            validatorKind = PdfExternalValidatorKind.Mustang;
            return true;
        }

        validatorKind = default;
        return false;
    }

    private bool HasPassingExternalValidation(PdfExternalValidatorKind validatorKind) {
        for (int i = 0; i < _externalValidations.Count; i++) {
            PdfExternalValidationResult result = _externalValidations[i];
            if (result.ValidatorKind == validatorKind &&
                IsExternalValidationForRequestedProfile(result) &&
                IsExternalValidationForArtifact(result) &&
                result.Status == PdfExternalValidationStatus.Passed) {
                return true;
            }
        }

        return false;
    }

    private PdfExternalValidatorProof BuildExternalValidatorProof(PdfExternalValidatorKind validatorKind) {
        var matches = new List<PdfExternalValidationResult>();
        for (int i = 0; i < _externalValidations.Count; i++) {
            PdfExternalValidationResult result = _externalValidations[i];
            if (result.ValidatorKind == validatorKind &&
                IsExternalValidationForRequestedProfile(result) &&
                IsExternalValidationForArtifact(result)) {
                matches.Add(result);
            }
        }

        return new PdfExternalValidatorProof(validatorKind, matches.AsReadOnly());
    }

    private int CountExternalValidations(PdfExternalValidationStatus status) {
        int count = 0;
        for (int i = 0; i < _externalValidations.Count; i++) {
            PdfExternalValidationResult result = _externalValidations[i];
            if (_requiredExternalValidators.Contains(result.ValidatorKind) &&
                IsExternalValidationForRequestedProfile(result) &&
                IsExternalValidationForArtifact(result) &&
                result.Status == status) {
                count++;
            }
        }

        return count;
    }

    private bool HasFailedExternalValidation(PdfExternalValidatorKind validatorKind) {
        for (int i = 0; i < _externalValidations.Count; i++) {
            PdfExternalValidationResult result = _externalValidations[i];
            if (result.ValidatorKind == validatorKind &&
                IsExternalValidationForRequestedProfile(result) &&
                IsExternalValidationForArtifact(result) &&
                (result.Status == PdfExternalValidationStatus.Failed ||
                 result.Status == PdfExternalValidationStatus.Error)) {
                return true;
            }
        }

        return false;
    }

    private bool IsExternalValidationForArtifact(PdfExternalValidationResult result) =>
        HasArtifactEvidence &&
        result.HasArtifactBinding &&
        string.Equals(result.ArtifactSha256, ArtifactSha256, StringComparison.OrdinalIgnoreCase) &&
        result.ArtifactSizeBytes == ArtifactSizeBytes;

    private bool IsExternalValidationForRequestedProfile(PdfExternalValidationResult result) {
        if (Profile == PdfComplianceProfile.None) {
            return string.IsNullOrWhiteSpace(result.Profile);
        }

        string? resultProfile = result.Profile;
        if (string.IsNullOrWhiteSpace(resultProfile)) {
            return true;
        }

        string normalizedResult = NormalizeProfileName(resultProfile!);
        foreach (string expectedProfileName in EnumerateExpectedExternalProfileNames(result.ValidatorKind)) {
            if (string.Equals(normalizedResult, NormalizeProfileName(expectedProfileName), StringComparison.OrdinalIgnoreCase)) {
                return true;
            }
        }

        return false;
    }

    private IEnumerable<string> EnumerateExpectedExternalProfileNames(PdfExternalValidatorKind validatorKind) {
        yield return DisplayName;
        yield return Profile.ToString();

        if ((Profile == PdfComplianceProfile.FacturX || Profile == PdfComplianceProfile.Zugferd) &&
            validatorKind == PdfExternalValidatorKind.VeraPdf) {
            yield return "PDF/A-3b";
            yield return PdfComplianceProfile.PdfA3B.ToString();
        }
    }

    private static string NormalizeProfileName(string profile) {
        var builder = new System.Text.StringBuilder(profile.Length);
        for (int i = 0; i < profile.Length; i++) {
            char ch = profile[i];
            if (char.IsLetterOrDigit(ch)) {
                builder.Append(char.ToUpperInvariant(ch));
            }
        }

        return builder.ToString();
    }
}
