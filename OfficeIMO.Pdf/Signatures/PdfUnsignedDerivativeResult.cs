namespace OfficeIMO.Pdf;

/// <summary>Full-rewrite PDF derivative with invalidated signature fields and revisions removed explicitly.</summary>
public sealed class PdfUnsignedDerivativeResult {
    private readonly byte[] _pdf;

    internal PdfUnsignedDerivativeResult(byte[] pdf, int removedSignatureCount) {
        _pdf = (byte[])pdf.Clone();
        RemovedSignatureCount = removedSignatureCount;
    }

    /// <summary>Unsigned, unencrypted full-rewrite derivative bytes.</summary>
    public byte[] Pdf => (byte[])_pdf.Clone();
    /// <summary>Signature definitions present in the source security inventory.</summary>
    public int RemovedSignatureCount { get; }
    /// <summary>Opens the derivative through the normal fluent API.</summary>
    public PdfDocument ToDocument() => PdfDocument.Load(_pdf);
}
