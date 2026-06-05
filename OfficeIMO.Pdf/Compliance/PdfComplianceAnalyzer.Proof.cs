namespace OfficeIMO.Pdf;

public static partial class PdfComplianceAnalyzer {
    /// <summary>
    /// Combines readiness diagnostics for the profile requested by the supplied options with external validator evidence.
    /// </summary>
    public static PdfComplianceProofReport AssessProof(PdfOptions options, IEnumerable<PdfExternalValidationResult>? externalValidations = null) {
        Guard.NotNull(options, nameof(options));
        return AssessProof(options.ComplianceProfile, options, externalValidations);
    }

    /// <summary>
    /// Combines readiness diagnostics for a requested profile with external validator evidence.
    /// </summary>
    public static PdfComplianceProofReport AssessProof(PdfComplianceProfile profile, PdfOptions options, IEnumerable<PdfExternalValidationResult>? externalValidations = null) {
        return AssessProof(profile, options, externalValidations, generatedStandardFonts: null);
    }

    /// <summary>
    /// Combines readiness diagnostics for a requested profile with generated-font evidence and external validator evidence.
    /// </summary>
    public static PdfComplianceProofReport AssessProof(PdfComplianceProfile profile, PdfOptions options, IEnumerable<PdfExternalValidationResult>? externalValidations, IEnumerable<PdfStandardFont>? generatedStandardFonts) {
        PdfComplianceReadinessReport readiness = Assess(profile, options, generatedStandardFonts);
        return AssessProof(readiness, externalValidations);
    }

    /// <summary>
    /// Combines an existing readiness report with external validator evidence.
    /// </summary>
    public static PdfComplianceProofReport AssessProof(PdfComplianceReadinessReport readiness, IEnumerable<PdfExternalValidationResult>? externalValidations = null) {
        Guard.NotNull(readiness, nameof(readiness));

        PdfExternalValidationResult[] validationSnapshot = SnapshotExternalValidations(externalValidations);
        return new PdfComplianceProofReport(
            readiness,
            GetRequiredExternalValidators(readiness.Profile),
            validationSnapshot);
    }

    private static PdfExternalValidationResult[] SnapshotExternalValidations(IEnumerable<PdfExternalValidationResult>? externalValidations) {
        if (externalValidations == null) {
            return Array.Empty<PdfExternalValidationResult>();
        }

        var snapshot = new List<PdfExternalValidationResult>();
        foreach (PdfExternalValidationResult result in externalValidations) {
            Guard.NotNull(result, nameof(externalValidations));
            snapshot.Add(result);
        }

        return snapshot.ToArray();
    }

    private static System.Collections.ObjectModel.ReadOnlyCollection<PdfExternalValidatorKind> GetRequiredExternalValidators(PdfComplianceProfile profile) {
        Guard.ComplianceProfile(profile, nameof(profile));

        var validators = new List<PdfExternalValidatorKind>();
        if (IsPdfA(profile) || IsElectronicInvoice(profile)) {
            validators.Add(PdfExternalValidatorKind.VeraPdf);
        }

        if (profile == PdfComplianceProfile.PdfUa1) {
            validators.Add(PdfExternalValidatorKind.PdfUaValidator);
        }

        if (IsElectronicInvoice(profile)) {
            validators.Add(PdfExternalValidatorKind.Mustang);
        }

        return validators.AsReadOnly();
    }
}
