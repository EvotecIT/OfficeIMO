namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>
    /// Assesses PDF/A and PDF/UA declarations found in XMP metadata against the exact PDF bytes.
    /// Declared metadata alone is never accepted as conformance proof.
    /// </summary>
    public PdfDeclaredComplianceClaimsReport AssessDeclaredComplianceClaims(
        IEnumerable<PdfExternalValidationResult>? externalValidations = null) {
        byte[] artifact = GetBytesForOperation();
        PdfReadDocument readDocument = PdfReadDocument.Open(artifact, ReadOptions);
        PdfDocumentInfo info = PdfInspector.FromReadDocument(readDocument, PdfInspector.Probe(artifact));
        PdfExternalValidationResult[] validations = externalValidations?.ToArray() ?? Array.Empty<PdfExternalValidationResult>();
        var claims = new List<PdfDeclaredComplianceClaim>(2);
        PdfXmpMetadataInfo? xmp = info.XmpMetadata;
        if (xmp?.HasPdfAIdentification == true) {
            claims.Add(BuildDeclaredPdfAClaim(xmp, artifact, validations));
        }
        if (xmp?.HasPdfUaIdentification == true) {
            claims.Add(BuildDeclaredPdfUaClaim(xmp, artifact, validations));
        }

        return new PdfDeclaredComplianceClaimsReport(
            PdfArtifactFingerprint.ComputeSha256(artifact),
            artifact.LongLength,
            claims.AsReadOnly());
    }

    private PdfDeclaredComplianceClaim BuildDeclaredPdfAClaim(
        PdfXmpMetadataInfo xmp,
        byte[] artifact,
        IReadOnlyList<PdfExternalValidationResult> validations) {
        int part = xmp.PdfAPart ?? 0;
        string conformance = string.IsNullOrWhiteSpace(xmp.PdfAConformance)
            ? string.Empty
            : xmp.PdfAConformance!.Trim().ToUpperInvariant();
        string declaration = "PDF/A-" + part.ToString(System.Globalization.CultureInfo.InvariantCulture) + conformance.ToLowerInvariant();
        PdfComplianceProfile? profile = MapPdfAProfile(part, conformance);
        return BuildDeclaredClaim(
            PdfDeclaredComplianceStandard.PdfA,
            declaration,
            profile,
            artifact,
            validations);
    }

    private PdfDeclaredComplianceClaim BuildDeclaredPdfUaClaim(
        PdfXmpMetadataInfo xmp,
        byte[] artifact,
        IReadOnlyList<PdfExternalValidationResult> validations) {
        int part = xmp.PdfUaPart ?? 0;
        string declaration = "PDF/UA-" + part.ToString(System.Globalization.CultureInfo.InvariantCulture);
        PdfComplianceProfile? profile = part switch {
            1 => PdfComplianceProfile.PdfUa1,
            2 => PdfComplianceProfile.PdfUa2,
            _ => null
        };
        return BuildDeclaredClaim(
            PdfDeclaredComplianceStandard.PdfUa,
            declaration,
            profile,
            artifact,
            validations);
    }

    private PdfDeclaredComplianceClaim BuildDeclaredClaim(
        PdfDeclaredComplianceStandard standard,
        string declaration,
        PdfComplianceProfile? profile,
        byte[] artifact,
        IReadOnlyList<PdfExternalValidationResult> validations) {
        if (!profile.HasValue) {
            return new PdfDeclaredComplianceClaim(
                standard,
                declaration,
                null,
                null,
                declaration + " is declared in XMP but is not mapped to an implemented OfficeIMO.Pdf compliance profile.");
        }

        PdfComplianceReadinessReport readiness = PdfComplianceAnalyzer.AssessReadback(profile.Value, artifact, ReadOptions);
        PdfComplianceProofReport proof = PdfComplianceAnalyzer.AssessProof(readiness, artifact, validations, ReadOptions);
        string diagnostic = proof.CanClaimConformance
            ? "Internal requirements and required external validators passed for the exact artifact."
            : proof.ProofStatus == "InternalGaps"
                ? "Known internal requirements are not satisfied for the declared profile."
                : proof.ExternalProofSummary;
        return new PdfDeclaredComplianceClaim(standard, declaration, profile, proof, diagnostic);
    }

    private static PdfComplianceProfile? MapPdfAProfile(int part, string conformance) {
        if (part == 2) {
            return conformance switch {
                "A" => PdfComplianceProfile.PdfA2A,
                "U" => PdfComplianceProfile.PdfA2U,
                "B" => PdfComplianceProfile.PdfA2B,
                _ => null
            };
        }
        if (part == 3) {
            return conformance switch {
                "A" => PdfComplianceProfile.PdfA3A,
                "U" => PdfComplianceProfile.PdfA3U,
                "B" => PdfComplianceProfile.PdfA3B,
                _ => null
            };
        }
        if (part == 4) {
            return conformance switch {
                "E" => PdfComplianceProfile.PdfA4E,
                "F" => PdfComplianceProfile.PdfA4F,
                "" => PdfComplianceProfile.PdfA4,
                _ => null
            };
        }
        return null;
    }
}
