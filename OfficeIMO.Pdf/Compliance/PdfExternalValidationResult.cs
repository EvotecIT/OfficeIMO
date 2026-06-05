namespace OfficeIMO.Pdf;

/// <summary>
/// Result from a caller- or CI-run external PDF compliance validator.
/// </summary>
public sealed class PdfExternalValidationResult {
    /// <summary>Creates an external validation result.</summary>
    public PdfExternalValidationResult(
        PdfExternalValidatorKind validatorKind,
        PdfExternalValidationStatus status,
        string validatorName,
        string diagnostic,
        string? profile = null,
        string? executablePath = null,
        string? arguments = null,
        int? exitCode = null) {
        ValidateValidatorKind(validatorKind, nameof(validatorKind));
        ValidateStatus(status, nameof(status));
        Guard.NotNullOrWhiteSpace(validatorName, nameof(validatorName));
        Guard.NotNullOrWhiteSpace(diagnostic, nameof(diagnostic));

        ValidatorKind = validatorKind;
        Status = status;
        ValidatorName = validatorName;
        Diagnostic = diagnostic;
        Profile = string.IsNullOrWhiteSpace(profile) ? null : profile;
        ExecutablePath = string.IsNullOrWhiteSpace(executablePath) ? null : executablePath;
        Arguments = string.IsNullOrWhiteSpace(arguments) ? null : arguments;
        ExitCode = exitCode;
    }

    /// <summary>Validator family.</summary>
    public PdfExternalValidatorKind ValidatorKind { get; }

    /// <summary>Validation outcome.</summary>
    public PdfExternalValidationStatus Status { get; }

    /// <summary>Human-readable validator name, for example veraPDF or Mustang.</summary>
    public string ValidatorName { get; }

    /// <summary>Human-readable result details from the caller or validator.</summary>
    public string Diagnostic { get; }

    /// <summary>Optional profile string used by the external tool.</summary>
    public string? Profile { get; }

    /// <summary>Optional executable path used by CI or wrapper tooling.</summary>
    public string? ExecutablePath { get; }

    /// <summary>Optional command-line arguments used by CI or wrapper tooling.</summary>
    public string? Arguments { get; }

    /// <summary>Optional process exit code reported by CI or wrapper tooling.</summary>
    public int? ExitCode { get; }

    /// <summary>Creates a passing validator result.</summary>
    public static PdfExternalValidationResult Passed(PdfExternalValidatorKind validatorKind, string validatorName, string diagnostic, string? profile = null) =>
        new(validatorKind, PdfExternalValidationStatus.Passed, validatorName, diagnostic, profile);

    /// <summary>Creates a failing validator result.</summary>
    public static PdfExternalValidationResult Failed(PdfExternalValidatorKind validatorKind, string validatorName, string diagnostic, string? profile = null, int? exitCode = null) =>
        new(validatorKind, PdfExternalValidationStatus.Failed, validatorName, diagnostic, profile, exitCode: exitCode);

    /// <summary>Creates a not-run validator result.</summary>
    public static PdfExternalValidationResult NotRun(PdfExternalValidatorKind validatorKind, string validatorName, string diagnostic, string? profile = null) =>
        new(validatorKind, PdfExternalValidationStatus.NotRun, validatorName, diagnostic, profile);

    /// <summary>Creates a validator error result.</summary>
    public static PdfExternalValidationResult Error(PdfExternalValidatorKind validatorKind, string validatorName, string diagnostic, string? profile = null, int? exitCode = null) =>
        new(validatorKind, PdfExternalValidationStatus.Error, validatorName, diagnostic, profile, exitCode: exitCode);

    private static void ValidateValidatorKind(PdfExternalValidatorKind value, string paramName) {
        if (value != PdfExternalValidatorKind.VeraPdf &&
            value != PdfExternalValidatorKind.PdfUaValidator &&
            value != PdfExternalValidatorKind.Mustang &&
            value != PdfExternalValidatorKind.Custom) {
            throw new System.ArgumentOutOfRangeException(paramName, "External validator kind must be VeraPdf, PdfUaValidator, Mustang, or Custom.");
        }
    }

    private static void ValidateStatus(PdfExternalValidationStatus value, string paramName) {
        if (value != PdfExternalValidationStatus.NotRun &&
            value != PdfExternalValidationStatus.Passed &&
            value != PdfExternalValidationStatus.Failed &&
            value != PdfExternalValidationStatus.Error) {
            throw new System.ArgumentOutOfRangeException(paramName, "External validation status must be NotRun, Passed, Failed, or Error.");
        }
    }
}
