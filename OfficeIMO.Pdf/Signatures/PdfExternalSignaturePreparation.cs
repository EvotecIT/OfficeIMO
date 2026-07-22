using System.Security.Cryptography;

namespace OfficeIMO.Pdf;

/// <summary>Prepared PDF bytes and byte ranges for an external signing operation.</summary>
public sealed class PdfExternalSignaturePreparation {
    private readonly byte[] _preparedPdf;
    private readonly PdfReadOptions? _readOptions;

    internal PdfExternalSignaturePreparation(
        byte[] preparedPdf,
        string fieldName,
        string filter,
        string subFilter,
        PdfSignatureProfile profile,
        IReadOnlyList<long> byteRangeValues,
        int contentsHexOffset,
        int contentsHexLength,
        int reservedSignatureContentsBytes,
        PdfReadOptions? readOptions = null) {
        _preparedPdf = (byte[])preparedPdf.Clone();
        _readOptions = readOptions;
        FieldName = fieldName;
        Filter = filter;
        SubFilter = subFilter;
        Profile = profile;
        ByteRangeValues = byteRangeValues.ToArray();
        ContentsHexOffset = contentsHexOffset;
        ContentsHexLength = contentsHexLength;
        ReservedSignatureContentsBytes = reservedSignatureContentsBytes;
    }

    /// <summary>PDF bytes containing a patched /ByteRange and zero-filled reserved /Contents placeholder.</summary>
    public byte[] PreparedPdf => (byte[])_preparedPdf.Clone();

    /// <summary>Name of the AcroForm signature field appended to the document.</summary>
    public string FieldName { get; }

    /// <summary>PDF signature filter emitted in the signature dictionary.</summary>
    public string Filter { get; }

    /// <summary>PDF signature subfilter emitted in the signature dictionary.</summary>
    public string SubFilter { get; }

    /// <summary>Approval, certification, or document-timestamp profile prepared for signing.</summary>
    public PdfSignatureProfile Profile { get; }

    /// <summary>Four-value detached signature /ByteRange array.</summary>
    public IReadOnlyList<long> ByteRangeValues { get; }

    /// <summary>Offset of the first hex character inside the reserved /Contents value.</summary>
    public int ContentsHexOffset { get; }

    /// <summary>Number of hex characters reserved inside /Contents.</summary>
    public int ContentsHexLength { get; }

    /// <summary>Number of raw signature bytes that can be injected into /Contents.</summary>
    public int ReservedSignatureContentsBytes { get; }

    /// <summary>Bytes covered by the /ByteRange and intended for external digest/signing.</summary>
    public byte[] SignedContent {
        get {
            int firstLength = checked((int)ByteRangeValues[1]);
            int secondLength = checked((int)ByteRangeValues[3]);
            var result = new byte[checked(firstLength + secondLength)];
            Buffer.BlockCopy(_preparedPdf, checked((int)ByteRangeValues[0]), result, 0, firstLength);
            Buffer.BlockCopy(_preparedPdf, checked((int)ByteRangeValues[2]), result, firstLength, secondLength);
            return result;
        }
    }

    /// <summary>Computes the SHA-256 digest of <see cref="SignedContent"/> for external signing services.</summary>
    public byte[] ComputeSha256Digest() {
        using IncrementalHash sha256 = IncrementalHash.CreateHash(HashAlgorithmName.SHA256);
        sha256.AppendData(_preparedPdf, checked((int)ByteRangeValues[0]), checked((int)ByteRangeValues[1]));
        sha256.AppendData(_preparedPdf, checked((int)ByteRangeValues[2]), checked((int)ByteRangeValues[3]));
        return sha256.GetHashAndReset();
    }

    /// <summary>
    /// Completes this in-memory preparation with detached CMS or timestamp bytes.
    /// The original read policy is preserved, while the input budget is expanded only for bytes appended by preparation.
    /// </summary>
    public PdfDocument Complete(byte[] signatureContents, PdfReadOptions? readOptions = null) {
        Guard.NotNull(signatureContents, nameof(signatureContents));
        byte[] completedPdf = PdfIncrementalUpdater.ApplyExternalSignature(this, signatureContents);
        PdfReadOptions effectiveOptions = readOptions ?? GetCompletionReadOptions(completedPdf.LongLength);
        return PdfDocument.Open(completedPdf, effectiveOptions);
    }

    internal PdfReadOptions GetCompletionReadOptions(long completedLength) =>
        PdfReadOptions.WithMinimumInputBytes(_readOptions, Math.Max(_preparedPdf.LongLength, completedLength));
}
