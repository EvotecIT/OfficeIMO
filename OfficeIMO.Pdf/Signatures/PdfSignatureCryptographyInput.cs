namespace OfficeIMO.Pdf;

/// <summary>Exact signed bytes and signature container supplied to an optional cryptography provider.</summary>
public sealed class PdfSignatureCryptographyInput {
    internal PdfSignatureCryptographyInput(
        PdfSignatureInfo signature,
        byte[] signedContent,
        byte[] signatureContents,
        long documentLength,
        PdfDocumentDssInfo documentSecurityStore) {
        Signature = signature;
        SignedContent = signedContent;
        SignatureContents = signatureContents;
        DocumentLength = documentLength;
        DocumentSecurityStore = documentSecurityStore;
    }

    /// <summary>PDF signature dictionary and field metadata.</summary>
    public PdfSignatureInfo Signature { get; }

    /// <summary>Exact concatenation of every byte segment covered by the signature `/ByteRange`.</summary>
    public byte[] SignedContent { get; }

    /// <summary>Decoded signature `/Contents` bytes, including any reserved trailing padding.</summary>
    public byte[] SignatureContents { get; }

    /// <summary>Complete PDF byte length used to validate the byte ranges.</summary>
    public long DocumentLength { get; }

    /// <summary>Document Security Store markers available to provider-specific LTV policy.</summary>
    public PdfDocumentDssInfo DocumentSecurityStore { get; }
}
