namespace OfficeIMO.Pdf;

/// <summary>Completed externally signed PDF and signer callback evidence.</summary>
public sealed class PdfExternalSignatureCompletion {
    private readonly byte[] _pdf;
    private readonly PdfReadOptions? _readOptions;

    internal PdfExternalSignatureCompletion(
        byte[] pdf,
        PdfExternalSignaturePreparation preparation,
        string signerName,
        int signatureContentsLength,
        PdfReadOptions? readOptions = null) {
        _pdf = (byte[])pdf.Clone();
        _readOptions = readOptions;
        Preparation = preparation;
        SignerName = signerName;
        SignatureContentsLength = signatureContentsLength;
    }

    /// <summary>PDF bytes with the callback-produced signature applied.</summary>
    public byte[] Pdf => (byte[])_pdf.Clone();

    /// <summary>Placeholder and byte-range preparation used by the signer.</summary>
    public PdfExternalSignaturePreparation Preparation { get; }

    /// <summary>Signer callback implementation name.</summary>
    public string SignerName { get; }

    /// <summary>CMS, CAdES, or timestamp token byte count returned by the signer.</summary>
    public int SignatureContentsLength { get; }

    /// <summary>Opens the completed PDF through the normal fluent document API.</summary>
    public PdfDocument ToDocument() => PdfDocument.Open(_pdf, _readOptions);
}
