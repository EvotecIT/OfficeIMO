namespace OfficeIMO.Pdf;

internal static partial class PdfComplianceAnalyzer {
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
    /// Combines readiness diagnostics, exact PDF artifact identity, generated-font evidence, and external validator evidence.
    /// </summary>
    public static PdfComplianceProofReport AssessProof(PdfComplianceProfile profile, PdfOptions options, byte[] artifact, IEnumerable<PdfExternalValidationResult>? externalValidations, IEnumerable<PdfStandardFont>? generatedStandardFonts) {
        PdfComplianceReadinessReport readiness = Assess(profile, options, generatedStandardFonts);
        return AssessProof(readiness, artifact, externalValidations);
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

    /// <summary>
    /// Combines an existing readiness report with exact PDF artifact identity and external validator evidence.
    /// </summary>
    public static PdfComplianceProofReport AssessProof(PdfComplianceReadinessReport readiness, byte[] artifact, IEnumerable<PdfExternalValidationResult>? externalValidations = null) {
        Guard.NotNull(readiness, nameof(readiness));
        Guard.NotNull(artifact, nameof(artifact));

        PdfComplianceReadinessReport artifactReadiness = AssessReadback(readiness.Profile, artifact);
        PdfComplianceReadinessReport combinedReadiness = CombineReadiness(readiness, artifactReadiness);
        PdfExternalValidationResult[] validationSnapshot = SnapshotExternalValidations(externalValidations);
        return new PdfComplianceProofReport(
            combinedReadiness,
            GetRequiredExternalValidators(combinedReadiness.Profile),
            validationSnapshot,
            PdfArtifactFingerprint.ComputeSha256(artifact),
            artifact.LongLength);
    }

    private static PdfComplianceReadinessReport CombineReadiness(PdfComplianceReadinessReport generated, PdfComplianceReadinessReport artifact) {
        if (generated.Profile != artifact.Profile) {
            throw new System.ArgumentException("Generated and artifact readiness reports must target the same compliance profile.", nameof(artifact));
        }

        var requirements = new List<PdfComplianceRequirement>(generated.Requirements.Count + artifact.Requirements.Count);
        var requirementIndexes = new Dictionary<string, int>(StringComparer.Ordinal);
        AddOrReconcile(generated.Requirements);
        AddOrReconcile(artifact.Requirements);
        return new PdfComplianceReadinessReport(generated.Profile, generated.DisplayName, requirements.AsReadOnly());

        void AddOrReconcile(IReadOnlyList<PdfComplianceRequirement> source) {
            for (int i = 0; i < source.Count; i++) {
                PdfComplianceRequirement requirement = source[i];
                if (!requirementIndexes.TryGetValue(requirement.Id, out int existingIndex)) {
                    requirementIndexes.Add(requirement.Id, requirements.Count);
                    requirements.Add(requirement);
                    continue;
                }

                PdfComplianceRequirement existing = requirements[existingIndex];
                if (GetBlockingRank(requirement.Status) > GetBlockingRank(existing.Status)) {
                    requirements[existingIndex] = requirement;
                }
            }
        }
    }

    private static int GetBlockingRank(PdfComplianceRequirementStatus status) {
        return status switch {
            PdfComplianceRequirementStatus.Satisfied => 0,
            PdfComplianceRequirementStatus.Missing => 1,
            PdfComplianceRequirementStatus.Unsupported => 2,
            _ => throw new ArgumentOutOfRangeException(nameof(status), status, "Unsupported PDF compliance requirement status.")
        };
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

        if (profile == PdfComplianceProfile.PdfUa1 ||
            profile == PdfComplianceProfile.PdfUa2) {
            validators.Add(PdfExternalValidatorKind.PdfUaValidator);
        }

        if (IsElectronicInvoice(profile)) {
            validators.Add(PdfExternalValidatorKind.Mustang);
        }

        return validators.AsReadOnly();
    }
}
