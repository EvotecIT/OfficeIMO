namespace OfficeIMO.Pdf;

/// <summary>Append-only DSS/VRI enrichment output and its validation proofs.</summary>
public sealed class PdfLongTermValidationEnrichmentResult {
    internal PdfLongTermValidationEnrichmentResult(
        byte[] pdf,
        string vriKey,
        PdfLongTermValidationEvidence evidence,
        PdfSignatureValidationReport validationBefore,
        PdfSignatureValidationReport validationAfter,
        PdfSignatureMutationReport mutationReport,
        IReadOnlyList<int> certificateObjectNumbers,
        IReadOnlyList<int> ocspObjectNumbers,
        IReadOnlyList<int> crlObjectNumbers) {
        Pdf = pdf;
        VriKey = vriKey;
        Evidence = evidence;
        ValidationBefore = validationBefore;
        ValidationAfter = validationAfter;
        MutationReport = mutationReport;
        CertificateObjectNumbers = certificateObjectNumbers;
        OcspObjectNumbers = ocspObjectNumbers;
        CrlObjectNumbers = crlObjectNumbers;
    }

    /// <summary>PDF bytes with the appended DSS/VRI revision.</summary>
    public byte[] Pdf { get; }

    /// <summary>Uppercase SHA-1 VRI key derived from the signature's complete hexadecimal Contents value.</summary>
    public string VriKey { get; }

    /// <summary>Requested validation material.</summary>
    public PdfLongTermValidationEvidence Evidence { get; }

    /// <summary>Structural and cryptographic validation before enrichment.</summary>
    public PdfSignatureValidationReport ValidationBefore { get; }

    /// <summary>Structural and cryptographic validation after enrichment.</summary>
    public PdfSignatureValidationReport ValidationAfter { get; }

    /// <summary>Byte-prefix, revision-chain, signature-range, and permission proof.</summary>
    public PdfSignatureMutationReport MutationReport { get; }

    /// <summary>New DSS certificate stream object numbers.</summary>
    public IReadOnlyList<int> CertificateObjectNumbers { get; }

    /// <summary>New DSS OCSP stream object numbers.</summary>
    public IReadOnlyList<int> OcspObjectNumbers { get; }

    /// <summary>New DSS CRL stream object numbers.</summary>
    public IReadOnlyList<int> CrlObjectNumbers { get; }

    /// <summary>True when the appended artifact passed every required structural preservation proof.</summary>
    public bool IsVerifiedAppendOnlyEnrichment =>
        MutationReport.IsPreservedAppendOnlyMutation &&
        ValidationAfter.Security.DocumentSecurityStore.VriKeys.Contains(VriKey, StringComparer.Ordinal) &&
        CertificateObjectNumbers.All(ValidationAfter.Security.DocumentSecurityStore.CertificateObjectNumbers.Contains) &&
        CertificateObjectNumbers.All(ValidationAfter.Security.DocumentSecurityStore.VriCertificateObjectNumbers.Contains) &&
        OcspObjectNumbers.All(ValidationAfter.Security.DocumentSecurityStore.OcspObjectNumbers.Contains) &&
        OcspObjectNumbers.All(ValidationAfter.Security.DocumentSecurityStore.VriOcspObjectNumbers.Contains) &&
        CrlObjectNumbers.All(ValidationAfter.Security.DocumentSecurityStore.CrlObjectNumbers.Contains) &&
        CrlObjectNumbers.All(ValidationAfter.Security.DocumentSecurityStore.VriCrlObjectNumbers.Contains);
}
