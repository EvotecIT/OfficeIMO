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
        CryptographicTrustVerified = false;
        DigestVerified = false;
        CertificateChainVerified = false;
        RevocationChecked = false;
        TimestampValidationPerformed = false;
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
    public bool IsStructurallyValid => !HasErrors;

    /// <summary>True when the file exposes DSS/VRI evidence used by long-term validation workflows.</summary>
    public bool HasLongTermValidationEvidence => Security.HasLongTermValidationEvidence;

    /// <summary>True when mutation should preserve the original PDF by appending revisions.</summary>
    public bool RequiresAppendOnlyMutation => Security.RequiresAppendOnlyMutation;

    /// <summary>False because OfficeIMO.Pdf does not perform certificate-chain or cryptographic signature verification.</summary>
    public bool CryptographicTrustVerified { get; }

    /// <summary>False because OfficeIMO.Pdf does not recompute signed byte-range digests.</summary>
    public bool DigestVerified { get; }

    /// <summary>False because OfficeIMO.Pdf does not build certificate chains.</summary>
    public bool CertificateChainVerified { get; }

    /// <summary>False because OfficeIMO.Pdf does not perform OCSP/CRL revocation checks.</summary>
    public bool RevocationChecked { get; }

    /// <summary>False because OfficeIMO.Pdf does not validate RFC 3161 timestamps cryptographically.</summary>
    public bool TimestampValidationPerformed { get; }

    /// <summary>True when any readable signature declares an RFC 3161 document timestamp subfilter.</summary>
    public bool HasDocumentTimestampSignature => Signatures.Any(static signature => signature.Signature.IsDocumentTimestamp);
}
