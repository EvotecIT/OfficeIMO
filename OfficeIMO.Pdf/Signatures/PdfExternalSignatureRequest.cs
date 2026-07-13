namespace OfficeIMO.Pdf;

/// <summary>Prepared signed content and digest passed to caller-owned key infrastructure.</summary>
public sealed class PdfExternalSignatureRequest {
    private readonly byte[] _signedContent;
    private readonly byte[] _sha256Digest;

    internal PdfExternalSignatureRequest(PdfExternalSignaturePreparation preparation) {
        Preparation = preparation;
        _signedContent = preparation.SignedContent;
        _sha256Digest = preparation.ComputeSha256Digest();
    }

    /// <summary>Prepared PDF placeholder and signature metadata.</summary>
    public PdfExternalSignaturePreparation Preparation { get; }

    /// <summary>Exact concatenated PDF byte ranges to sign.</summary>
    public byte[] SignedContent => (byte[])_signedContent.Clone();

    /// <summary>SHA-256 digest of <see cref="SignedContent"/> for digest-based remote signers.</summary>
    public byte[] Sha256Digest => (byte[])_sha256Digest.Clone();

    /// <summary>Digest algorithm name corresponding to <see cref="Sha256Digest"/>.</summary>
    public string DigestAlgorithmName { get; } = "SHA-256";
}
