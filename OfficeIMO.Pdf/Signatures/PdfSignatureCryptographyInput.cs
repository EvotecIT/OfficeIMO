namespace OfficeIMO.Pdf;

/// <summary>Exact signed bytes and signature container supplied to an optional cryptography provider.</summary>
public sealed class PdfSignatureCryptographyInput {
    private readonly byte[] _signedContent;
    private readonly byte[] _signatureContents;

    internal PdfSignatureCryptographyInput(
        PdfSignatureInfo signature,
        byte[] signedContent,
        byte[] signatureContents,
        long documentLength,
        PdfDocumentDssInfo documentSecurityStore,
        PdfDocumentValidationData? validationData = null) {
        Signature = signature;
        _signedContent = (byte[])signedContent.Clone();
        _signatureContents = (byte[])signatureContents.Clone();
        DocumentLength = documentLength;
        DocumentSecurityStore = documentSecurityStore;
        ValidationData = validationData ?? PdfDocumentValidationData.Empty;
    }

    /// <summary>Decoded DSS certificates, OCSP responses and CRLs, with proof-of-existence times from document time-stamps.</summary>
    internal PdfDocumentValidationData ValidationData { get; }

    /// <summary>PDF signature dictionary and field metadata.</summary>
    public PdfSignatureInfo Signature { get; }

    /// <summary>Exact concatenation of every byte segment covered by the signature `/ByteRange`.</summary>
    public byte[] SignedContent => (byte[])_signedContent.Clone();

    /// <summary>Decoded signature `/Contents` bytes, including any reserved trailing padding.</summary>
    public byte[] SignatureContents => (byte[])_signatureContents.Clone();

    /// <summary>Complete PDF byte length used to validate the byte ranges.</summary>
    public long DocumentLength { get; }

    /// <summary>Document Security Store markers available to provider-specific LTV policy.</summary>
    public PdfDocumentDssInfo DocumentSecurityStore { get; }
}
