namespace OfficeIMO.Pdf;

/// <summary>
/// Machine-readable proof row for one external validator family required by a compliance profile.
/// </summary>
public sealed class PdfExternalValidatorProof {
    internal PdfExternalValidatorProof(
        PdfExternalValidatorKind validatorKind,
        IReadOnlyList<PdfExternalValidationResult> validations) {
        Guard.NotNull(validations, nameof(validations));

        ValidatorKind = validatorKind;
        Validations = validations;
    }

    /// <summary>Required validator family represented by this proof row.</summary>
    public PdfExternalValidatorKind ValidatorKind { get; }

    /// <summary>Matching validation results for the requested profile and validator family.</summary>
    public IReadOnlyList<PdfExternalValidationResult> Validations { get; }

    /// <summary>First matching passing validation result, when supplied.</summary>
    public PdfExternalValidationResult? PassingValidation => Find(PdfExternalValidationStatus.Passed);

    /// <summary>First matching failed or errored validation result, when supplied.</summary>
    public PdfExternalValidationResult? BlockingValidation =>
        Find(PdfExternalValidationStatus.Failed) ?? Find(PdfExternalValidationStatus.Error);

    /// <summary>Most relevant validation result for display and automation.</summary>
    public PdfExternalValidationResult? PrimaryValidation =>
        BlockingValidation ?? PassingValidation ?? Find(PdfExternalValidationStatus.NotRun);

    /// <summary>Stable row status: Missing, NotRun, Passed, Failed, or Error.</summary>
    public PdfExternalValidatorProofStatus Status {
        get {
            PdfExternalValidationResult? failed = Find(PdfExternalValidationStatus.Failed);
            if (failed is not null) {
                return PdfExternalValidatorProofStatus.Failed;
            }

            PdfExternalValidationResult? error = Find(PdfExternalValidationStatus.Error);
            if (error is not null) {
                return PdfExternalValidatorProofStatus.Error;
            }

            if (PassingValidation is not null) {
                return PdfExternalValidatorProofStatus.Passed;
            }

            if (Find(PdfExternalValidationStatus.NotRun) is not null) {
                return PdfExternalValidatorProofStatus.NotRun;
            }

            return PdfExternalValidatorProofStatus.Missing;
        }
    }

    /// <summary>True when the required validator supplied a passing result and no matching failure or error exists.</summary>
    public bool IsSatisfied => Status == PdfExternalValidatorProofStatus.Passed;

    /// <summary>True when the required validator has not supplied a passing result.</summary>
    public bool IsMissing => Status == PdfExternalValidatorProofStatus.Missing || Status == PdfExternalValidatorProofStatus.NotRun;

    /// <summary>True when the validator supplied a failed or errored result that blocks a conformance claim.</summary>
    public bool HasBlockingValidation =>
        Status == PdfExternalValidatorProofStatus.Failed || Status == PdfExternalValidatorProofStatus.Error;

    /// <summary>True when this validator row prevents claiming conformance.</summary>
    public bool BlocksConformanceClaim => !IsSatisfied;

    /// <summary>Display name from the primary validation result, or the validator family name when no result was supplied.</summary>
    public string ValidatorName => PrimaryValidation?.ValidatorName ?? ValidatorKind.ToString();

    /// <summary>Diagnostic from the primary validation result, or a missing-evidence message when no result was supplied.</summary>
    public string Diagnostic => PrimaryValidation?.Diagnostic ?? "Missing external validation.";

    /// <summary>Profile string reported by the primary validation result, when supplied.</summary>
    public string? Profile => PrimaryValidation?.Profile;

    /// <summary>Process exit code reported by the primary validation result, when supplied.</summary>
    public int? ExitCode => PrimaryValidation?.ExitCode;

    private PdfExternalValidationResult? Find(PdfExternalValidationStatus status) {
        for (int i = 0; i < Validations.Count; i++) {
            PdfExternalValidationResult validation = Validations[i];
            if (validation.Status == status) {
                return validation;
            }
        }

        return null;
    }
}
