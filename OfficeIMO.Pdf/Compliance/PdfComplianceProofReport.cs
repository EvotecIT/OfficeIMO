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
        IReadOnlyList<PdfExternalValidationResult> externalValidations) {
        Guard.NotNull(readiness, nameof(readiness));
        Guard.NotNull(requiredExternalValidators, nameof(requiredExternalValidators));
        Guard.NotNull(externalValidations, nameof(externalValidations));

        Readiness = readiness;
        _requiredExternalValidators = requiredExternalValidators;
        _externalValidations = externalValidations;
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
        IsInternallyReady &&
        HasRequiredExternalValidation &&
        FailedExternalValidations.Count == 0;

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
                    !IsExternalValidationForRequestedProfile(result)) {
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
                IsExternalValidationForRequestedProfile(result)) {
                return result;
            }
        }

        return null;
    }

    internal static bool IsExternalValidationRequirement(string id) =>
        string.Equals(id, "verapdf-validation", StringComparison.Ordinal) ||
        string.Equals(id, "pdfua-validation", StringComparison.Ordinal) ||
        string.Equals(id, "mustang-validation", StringComparison.Ordinal);

    private bool HasPassingExternalValidation(PdfExternalValidatorKind validatorKind) {
        for (int i = 0; i < _externalValidations.Count; i++) {
            PdfExternalValidationResult result = _externalValidations[i];
            if (result.ValidatorKind == validatorKind &&
                IsExternalValidationForRequestedProfile(result) &&
                result.Status == PdfExternalValidationStatus.Passed) {
                return true;
            }
        }

        return false;
    }

    private bool HasFailedExternalValidation(PdfExternalValidatorKind validatorKind) {
        for (int i = 0; i < _externalValidations.Count; i++) {
            PdfExternalValidationResult result = _externalValidations[i];
            if (result.ValidatorKind == validatorKind &&
                IsExternalValidationForRequestedProfile(result) &&
                (result.Status == PdfExternalValidationStatus.Failed ||
                 result.Status == PdfExternalValidationStatus.Error)) {
                return true;
            }
        }

        return false;
    }

    private bool IsExternalValidationForRequestedProfile(PdfExternalValidationResult result) {
        if (Profile == PdfComplianceProfile.None) {
            return string.IsNullOrWhiteSpace(result.Profile);
        }

        string? resultProfile = result.Profile;
        if (string.IsNullOrWhiteSpace(resultProfile)) {
            return false;
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
