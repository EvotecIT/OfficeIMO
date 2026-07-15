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
        int? exitCode = null,
        string? validatorVersion = null,
        string? artifactSha256 = null,
        long? artifactSizeBytes = null,
        System.DateTimeOffset? validatedAtUtc = null,
        IEnumerable<string>? warnings = null) {
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
        ValidatorVersion = string.IsNullOrWhiteSpace(validatorVersion) ? null : validatorVersion!.Trim();
        ArtifactSha256 = PdfArtifactFingerprint.NormalizeSha256(artifactSha256, nameof(artifactSha256));
        if (artifactSizeBytes.HasValue && artifactSizeBytes.Value < 0) {
            throw new System.ArgumentOutOfRangeException(nameof(artifactSizeBytes), "Artifact size cannot be negative.");
        }

        ArtifactSizeBytes = artifactSizeBytes;
        ValidatedAtUtc = validatedAtUtc?.ToUniversalTime();
        Warnings = SnapshotWarnings(warnings);
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

    /// <summary>Validator version reported by the proof runner.</summary>
    public string? ValidatorVersion { get; }

    /// <summary>Lowercase SHA-256 of the exact PDF bytes supplied to the validator.</summary>
    public string? ArtifactSha256 { get; }

    /// <summary>Size of the exact validated PDF artifact, in bytes.</summary>
    public long? ArtifactSizeBytes { get; }

    /// <summary>UTC timestamp recorded by the proof runner after validation.</summary>
    public System.DateTimeOffset? ValidatedAtUtc { get; }

    /// <summary>Warnings reported by the validator or proof runner.</summary>
    public IReadOnlyList<string> Warnings { get; }

    /// <summary>True when this result identifies an exact PDF artifact by SHA-256 and byte length.</summary>
    public bool HasArtifactBinding => ArtifactSha256 != null && ArtifactSizeBytes.HasValue;

    /// <summary>Creates a passing validator result.</summary>
    public static PdfExternalValidationResult Passed(PdfExternalValidatorKind validatorKind, string validatorName, string diagnostic, string? profile = null) =>
        new(validatorKind, PdfExternalValidationStatus.Passed, validatorName, diagnostic, profile);

    /// <summary>Creates a passing result bound to the exact supplied PDF bytes.</summary>
    public static PdfExternalValidationResult PassedForArtifact(
        PdfExternalValidatorKind validatorKind,
        string validatorName,
        string validatorVersion,
        string diagnostic,
        byte[] artifact,
        string? profile = null,
        IEnumerable<string>? warnings = null,
        string? executablePath = null,
        string? arguments = null,
        int? exitCode = 0,
        System.DateTimeOffset? validatedAtUtc = null) {
        Guard.NotNull(artifact, nameof(artifact));
        return new PdfExternalValidationResult(
            validatorKind,
            PdfExternalValidationStatus.Passed,
            validatorName,
            diagnostic,
            profile,
            executablePath,
            arguments,
            exitCode,
            validatorVersion,
            PdfArtifactFingerprint.ComputeSha256(artifact),
            artifact.LongLength,
            validatedAtUtc ?? System.DateTimeOffset.UtcNow,
            warnings);
    }

    /// <summary>Creates a failing validator result.</summary>
    public static PdfExternalValidationResult Failed(PdfExternalValidatorKind validatorKind, string validatorName, string diagnostic, string? profile = null, int? exitCode = null) =>
        new(validatorKind, PdfExternalValidationStatus.Failed, validatorName, diagnostic, profile, exitCode: exitCode);

    /// <summary>Creates a not-run validator result.</summary>
    public static PdfExternalValidationResult NotRun(PdfExternalValidatorKind validatorKind, string validatorName, string diagnostic, string? profile = null) =>
        new(validatorKind, PdfExternalValidationStatus.NotRun, validatorName, diagnostic, profile);

    /// <summary>Creates a not-run result associated with an intended exact PDF artifact.</summary>
    public static PdfExternalValidationResult NotRunForArtifact(
        PdfExternalValidatorKind validatorKind,
        string validatorName,
        string validatorVersion,
        string diagnostic,
        byte[] artifact,
        string? profile = null,
        IEnumerable<string>? warnings = null,
        System.DateTimeOffset? recordedAtUtc = null) {
        Guard.NotNull(artifact, nameof(artifact));
        return new PdfExternalValidationResult(
            validatorKind,
            PdfExternalValidationStatus.NotRun,
            validatorName,
            diagnostic,
            profile,
            validatorVersion: validatorVersion,
            artifactSha256: PdfArtifactFingerprint.ComputeSha256(artifact),
            artifactSizeBytes: artifact.LongLength,
            validatedAtUtc: recordedAtUtc ?? System.DateTimeOffset.UtcNow,
            warnings: warnings);
    }

    /// <summary>Creates a validator error result.</summary>
    public static PdfExternalValidationResult Error(PdfExternalValidatorKind validatorKind, string validatorName, string diagnostic, string? profile = null, int? exitCode = null) =>
        new(validatorKind, PdfExternalValidationStatus.Error, validatorName, diagnostic, profile, exitCode: exitCode);

    /// <summary>Creates a validator result from a process exit code, treating the configured success code as Passed and all other codes as Failed.</summary>
    public static PdfExternalValidationResult FromExitCode(
        PdfExternalValidatorKind validatorKind,
        int exitCode,
        string validatorName,
        string diagnostic,
        string? profile = null,
        string? executablePath = null,
        string? arguments = null,
        int successExitCode = 0) =>
        new(
            validatorKind,
            exitCode == successExitCode ? PdfExternalValidationStatus.Passed : PdfExternalValidationStatus.Failed,
            validatorName,
            diagnostic,
            profile,
            executablePath,
            arguments,
            exitCode);

    /// <summary>Creates an exit-code result bound to the exact supplied PDF bytes.</summary>
    public static PdfExternalValidationResult FromExitCodeForArtifact(
        PdfExternalValidatorKind validatorKind,
        int exitCode,
        string validatorName,
        string validatorVersion,
        string diagnostic,
        byte[] artifact,
        string? profile = null,
        string? executablePath = null,
        string? arguments = null,
        int successExitCode = 0,
        IEnumerable<string>? warnings = null,
        System.DateTimeOffset? validatedAtUtc = null) {
        Guard.NotNull(artifact, nameof(artifact));
        return new PdfExternalValidationResult(
            validatorKind,
            exitCode == successExitCode ? PdfExternalValidationStatus.Passed : PdfExternalValidationStatus.Failed,
            validatorName,
            diagnostic,
            profile,
            executablePath,
            arguments,
            exitCode,
            validatorVersion,
            PdfArtifactFingerprint.ComputeSha256(artifact),
            artifact.LongLength,
            validatedAtUtc ?? System.DateTimeOffset.UtcNow,
            warnings);
    }

    private static IReadOnlyList<string> SnapshotWarnings(IEnumerable<string>? warnings) {
        if (warnings == null) {
            return Array.Empty<string>();
        }

        var snapshot = new List<string>();
        foreach (string warning in warnings) {
            if (!string.IsNullOrWhiteSpace(warning)) {
                snapshot.Add(warning.Trim());
            }
        }

        return snapshot.AsReadOnly();
    }

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
