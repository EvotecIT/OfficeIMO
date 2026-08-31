namespace OfficeIMO.Pdf;

/// <summary>Standards whose declarations can be recognized from PDF XMP metadata.</summary>
public enum PdfDeclaredComplianceStandard {
    /// <summary>PDF/A archival conformance declaration.</summary>
    PdfA,
    /// <summary>PDF/UA accessibility conformance declaration.</summary>
    PdfUa
}

/// <summary>Stable status for one declared compliance claim assessed against exact artifact evidence.</summary>
public enum PdfDeclaredComplianceClaimStatus {
    /// <summary>The declaration cannot be mapped to a profile implemented by this OfficeIMO.Pdf version.</summary>
    UnsupportedProfile,
    /// <summary>Known internal requirements are not satisfied.</summary>
    InternalGaps,
    /// <summary>The exact artifact has not been bound to the proof report.</summary>
    MissingArtifactEvidence,
    /// <summary>Required external validation for the exact artifact is missing.</summary>
    MissingExternalValidation,
    /// <summary>A required external validator failed or errored for the exact artifact.</summary>
    ExternalValidationFailed,
    /// <summary>Internal requirements and every required external validator passed for the exact artifact.</summary>
    Claimable
}

/// <summary>Assessment of one PDF/A or PDF/UA declaration found in XMP metadata.</summary>
public sealed class PdfDeclaredComplianceClaim {
    internal PdfDeclaredComplianceClaim(
        PdfDeclaredComplianceStandard standard,
        string declaration,
        PdfComplianceProfile? profile,
        PdfComplianceProofReport? proof,
        string diagnostic) {
        Standard = standard;
        Declaration = declaration;
        Profile = profile;
        Proof = proof;
        Diagnostic = diagnostic;
        Status = ResolveStatus(profile, proof);
    }

    /// <summary>Declared standard family.</summary>
    public PdfDeclaredComplianceStandard Standard { get; }

    /// <summary>Normalized declaration read from XMP metadata, such as PDF/A-3b or PDF/UA-1.</summary>
    public string Declaration { get; }

    /// <summary>Mapped OfficeIMO.Pdf profile, or null when the declared profile is not implemented.</summary>
    public PdfComplianceProfile? Profile { get; }

    /// <summary>Exact-artifact proof report for recognized profiles.</summary>
    public PdfComplianceProofReport? Proof { get; }

    /// <summary>Stable claim status.</summary>
    public PdfDeclaredComplianceClaimStatus Status { get; }

    /// <summary>Human-readable reason for the current status.</summary>
    public string Diagnostic { get; }

    /// <summary>True when the declaration maps to a supported OfficeIMO.Pdf profile.</summary>
    public bool IsRecognized => Profile.HasValue && Proof is not null;

    /// <summary>True only when internal checks and required external validation passed for the exact artifact.</summary>
    public bool CanClaimConformance => Status == PdfDeclaredComplianceClaimStatus.Claimable;

    private static PdfDeclaredComplianceClaimStatus ResolveStatus(
        PdfComplianceProfile? profile,
        PdfComplianceProofReport? proof) {
        if (!profile.HasValue || proof is null) return PdfDeclaredComplianceClaimStatus.UnsupportedProfile;
        return proof.ProofStatus switch {
            "InternalGaps" => PdfDeclaredComplianceClaimStatus.InternalGaps,
            "MissingArtifactEvidence" => PdfDeclaredComplianceClaimStatus.MissingArtifactEvidence,
            "MissingExternalValidation" => PdfDeclaredComplianceClaimStatus.MissingExternalValidation,
            "ExternalValidationFailed" => PdfDeclaredComplianceClaimStatus.ExternalValidationFailed,
            "Claimable" => PdfDeclaredComplianceClaimStatus.Claimable,
            _ => PdfDeclaredComplianceClaimStatus.InternalGaps
        };
    }
}

/// <summary>All PDF/A and PDF/UA declarations discovered and assessed for one exact PDF artifact.</summary>
public sealed class PdfDeclaredComplianceClaimsReport {
    internal PdfDeclaredComplianceClaimsReport(
        string artifactSha256,
        long artifactSizeBytes,
        IReadOnlyList<PdfDeclaredComplianceClaim> claims) {
        ArtifactSha256 = artifactSha256;
        ArtifactSizeBytes = artifactSizeBytes;
        Claims = claims;
    }

    /// <summary>SHA256 of the exact PDF artifact that was inspected.</summary>
    public string ArtifactSha256 { get; }

    /// <summary>Size of the exact PDF artifact that was inspected.</summary>
    public long ArtifactSizeBytes { get; }

    /// <summary>PDF/A and PDF/UA declarations in stable standard order.</summary>
    public IReadOnlyList<PdfDeclaredComplianceClaim> Claims { get; }

    /// <summary>True when at least one PDF/A or PDF/UA declaration was found, including an unsupported declaration.</summary>
    public bool HasClaims => Claims.Count > 0;

    /// <summary>Claims that map to supported OfficeIMO.Pdf profiles.</summary>
    public IReadOnlyList<PdfDeclaredComplianceClaim> RecognizedClaims =>
        Claims.Where(static claim => claim.IsRecognized).ToArray();

    /// <summary>Declarations that cannot be mapped to a supported profile.</summary>
    public IReadOnlyList<PdfDeclaredComplianceClaim> UnsupportedClaims =>
        Claims.Where(static claim => !claim.IsRecognized).ToArray();

    /// <summary>
    /// True only when at least one declaration exists and every declaration is recognized and claimable
    /// for the exact artifact. A metadata declaration by itself never satisfies this gate.
    /// </summary>
    public bool CanClaimAllDeclaredConformance =>
        Claims.Count > 0 && Claims.All(static claim => claim.CanClaimConformance);
}
