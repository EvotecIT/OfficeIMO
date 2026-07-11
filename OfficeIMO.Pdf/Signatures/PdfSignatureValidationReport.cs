namespace OfficeIMO.Pdf;

/// <summary>Dependency-free PDF signature validation and preservation report.</summary>
public sealed class PdfSignatureValidationReport {
    internal PdfSignatureValidationReport(
        PdfDocumentSecurityInfo security,
        long fileLength,
        IReadOnlyList<PdfSignatureValidationResult> signatures,
        IReadOnlyList<PdfSignatureValidationFinding> findings,
        bool objectGraphParsed,
        string? objectGraphError) {
        Security = security;
        FileLength = fileLength;
        Signatures = signatures;
        Findings = findings;
        ObjectGraphParsed = objectGraphParsed;
        ObjectGraphError = objectGraphError;
    }

    /// <summary>Security, signature, and revision markers read from the PDF.</summary>
    public PdfDocumentSecurityInfo Security { get; }

    /// <summary>Input PDF length in bytes.</summary>
    public long FileLength { get; }

    /// <summary>Per-signature structural validation results.</summary>
    public IReadOnlyList<PdfSignatureValidationResult> Signatures { get; }

    /// <summary>All validation findings.</summary>
    public IReadOnlyList<PdfSignatureValidationFinding> Findings { get; }

    /// <summary>True when the object graph was parsed enough to inspect signature dictionaries.</summary>
    public bool ObjectGraphParsed { get; }

    /// <summary>Object graph parse failure, when signature validation was limited.</summary>
    public string? ObjectGraphError { get; }

    /// <summary>True when signature fields or signature values were detected.</summary>
    public bool HasSignatures => Security.HasSignatures;

    /// <summary>Number of readable signature validation results.</summary>
    public int SignatureCount => Signatures.Count;

    /// <summary>True when any validation finding is an error.</summary>
    public bool HasErrors => Findings.Any(static finding => finding.Severity == PdfDiagnosticSeverity.Error);

    /// <summary>True when any validation finding is a warning.</summary>
    public bool HasWarnings => Findings.Any(static finding => finding.Severity == PdfDiagnosticSeverity.Warning);

    /// <summary>True when no structural validation errors were found.</summary>
    public bool IsStructurallyValid =>
        ObjectGraphParsed &&
        Findings.All(static finding => finding.IsCryptographic || finding.Severity != PdfDiagnosticSeverity.Error);

    /// <summary>True when the file exposes DSS/VRI evidence used by long-term validation workflows.</summary>
    public bool HasLongTermValidationEvidence => Security.HasLongTermValidationEvidence;

    /// <summary>True when mutation should preserve the original PDF by appending revisions.</summary>
    public bool RequiresAppendOnlyMutation => Security.RequiresAppendOnlyMutation;

    /// <summary>True when every readable signature has a provider result.</summary>
    public bool CryptographicValidationPerformed =>
        Signatures.Count > 0 && Signatures.All(static signature => signature.CryptographicResult is not null);

    /// <summary>True when every readable signature passed signature math, digest, and certificate-chain validation.</summary>
    public bool CryptographicTrustVerified =>
        CryptographicValidationPerformed &&
        Signatures.All(static signature =>
            signature.CryptographicResult!.IsMathematicallyValid &&
            signature.CryptographicResult.CertificateChainStatus == PdfCryptographicValidationStatus.Valid);

    /// <summary>True when every provider verified its signature's signed-content digest.</summary>
    public bool DigestVerified =>
        CryptographicValidationPerformed &&
        Signatures.All(static signature => signature.CryptographicResult!.MessageDigestStatus == PdfCryptographicValidationStatus.Valid);

    /// <summary>True when every provider reported a valid signer or TSA certificate chain.</summary>
    public bool CertificateChainVerified =>
        CryptographicValidationPerformed &&
        Signatures.All(static signature => signature.CryptographicResult!.CertificateChainStatus == PdfCryptographicValidationStatus.Valid);

    /// <summary>True when every provider performed a definitive revocation check, whether valid or revoked.</summary>
    public bool RevocationChecked =>
        CryptographicValidationPerformed &&
        Signatures.All(static signature =>
            signature.CryptographicResult!.RevocationStatus == PdfCryptographicValidationStatus.Valid ||
            signature.CryptographicResult.RevocationStatus == PdfCryptographicValidationStatus.Invalid);

    /// <summary>True when each timestamp-bearing signature received a definitive timestamp result.</summary>
    public bool TimestampValidationPerformed =>
        CryptographicValidationPerformed &&
        Signatures.Any(static signature =>
            signature.Signature.IsDocumentTimestamp ||
            signature.CryptographicResult!.TimestampStatus != PdfCryptographicValidationStatus.NotPerformed) &&
        Signatures
            .Where(static signature =>
                signature.Signature.IsDocumentTimestamp ||
                signature.CryptographicResult!.TimestampStatus != PdfCryptographicValidationStatus.NotPerformed)
            .All(static signature =>
                signature.CryptographicResult!.TimestampStatus == PdfCryptographicValidationStatus.Valid ||
                signature.CryptographicResult.TimestampStatus == PdfCryptographicValidationStatus.Invalid);

    /// <summary>True when every provider validated public-key signature math and signed-content digest.</summary>
    public bool MathematicalSignaturesVerified =>
        CryptographicValidationPerformed &&
        Signatures.All(static signature => signature.CryptographicResult!.IsMathematicallyValid);

    /// <summary>True when any readable signature declares an RFC 3161 document timestamp subfilter.</summary>
    public bool HasDocumentTimestampSignature => Signatures.Any(static signature => signature.Signature.IsDocumentTimestamp);

    /// <summary>True when the file exposes enough structural LTV markers for an external validator to attempt offline long-term validation.</summary>
    public bool HasOfflineLongTermValidationReadiness =>
        HasLongTermValidationEvidence &&
        Security.DocumentSecurityStore.TopLevelEvidenceObjectCount > 0 &&
        Security.DocumentSecurityStore.VriEntryCount > 0;

    /// <summary>Stable no-dependency signature proof state: Unsigned, StructuralIssues, LtvEvidenceReady, LtvEvidencePartial, or ExternalCryptoValidationRequired.</summary>
    public string ProofStatus {
        get {
            if (!HasSignatures) {
                return "Unsigned";
            }

            if (!IsStructurallyValid) {
                return "StructuralIssues";
            }

            if (CryptographicValidationPerformed) {
                if (!MathematicalSignaturesVerified) {
                    return "CryptographicInvalid";
                }

                return CryptographicTrustVerified
                    ? "CryptographicallyValidAndTrusted"
                    : "CryptographicallyValidTrustIndeterminate";
            }

            if (HasOfflineLongTermValidationReadiness) {
                return "LtvEvidenceReady";
            }

            if (HasLongTermValidationEvidence) {
                return "LtvEvidencePartial";
            }

            return "ExternalCryptoValidationRequired";
        }
    }
}
