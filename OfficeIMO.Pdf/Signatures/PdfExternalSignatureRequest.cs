namespace OfficeIMO.Pdf;

/// <summary>Prepared signed content and digest passed to caller-owned key infrastructure.</summary>
public sealed class PdfExternalSignatureRequest {
    internal PdfExternalSignatureRequest(PdfExternalSignaturePreparation preparation) {
        Preparation = preparation;
        SignedContent = preparation.SignedContent;
        Sha256Digest = preparation.ComputeSha256Digest();
    }

    /// <summary>Prepared PDF placeholder and signature metadata.</summary>
    public PdfExternalSignaturePreparation Preparation { get; }

    /// <summary>Exact concatenated PDF byte ranges to sign.</summary>
    public byte[] SignedContent { get; }

    /// <summary>SHA-256 digest of <see cref="SignedContent"/> for digest-based remote signers.</summary>
    public byte[] Sha256Digest { get; }

    /// <summary>Digest algorithm name corresponding to <see cref="Sha256Digest"/>.</summary>
    public string DigestAlgorithmName { get; } = "SHA-256";
}
