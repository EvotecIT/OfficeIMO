namespace OfficeIMO.Pdf;

/// <summary>
/// Lightweight metadata read from a PDF signature value dictionary and its owning AcroForm signature field.
/// </summary>
public sealed class PdfSignatureInfo {
    internal PdfSignatureInfo(
        int objectNumber,
        int? fieldObjectNumber,
        string? fieldName,
        PdfSignatureFieldLockInfo? fieldLock,
        PdfSignatureSeedValueInfo? seedValue,
        string? filter,
        string? subFilter,
        string? signerName,
        string? location,
        string? reason,
        string? contactInfo,
        string? signingTimeRaw,
        bool hasByteRange,
        IReadOnlyList<long> byteRangeValues,
        int byteRangeValueCount,
        bool hasContents,
        byte[]? contentsBytes,
        int? contentsSizeBytes,
        int? contentsEncodedSizeBytes,
        int referenceCount) {
        ObjectNumber = objectNumber;
        FieldObjectNumber = fieldObjectNumber;
        FieldName = fieldName;
        FieldLock = fieldLock;
        SeedValue = seedValue;
        Filter = filter;
        SubFilter = subFilter;
        SignerName = signerName;
        Location = location;
        Reason = reason;
        ContactInfo = contactInfo;
        SigningTimeRaw = signingTimeRaw;
        HasByteRange = hasByteRange;
        ByteRangeValues = byteRangeValues;
        ByteRangeValueCount = byteRangeValueCount;
        HasContents = hasContents;
        ContentsBytes = contentsBytes;
        ContentsSizeBytes = contentsSizeBytes;
        ContentsEncodedSizeBytes = contentsEncodedSizeBytes;
        ReferenceCount = referenceCount;
    }

    /// <summary>Object number of the signature value dictionary.</summary>
    public int ObjectNumber { get; }

    /// <summary>Object number of the AcroForm signature field whose /V points to this signature value, when found.</summary>
    public int? FieldObjectNumber { get; }

    /// <summary>Readable AcroForm signature field name, when found.</summary>
    public string? FieldName { get; }

    /// <summary>Signature field /Lock constraints, when present.</summary>
    public PdfSignatureFieldLockInfo? FieldLock { get; }

    /// <summary>Signature field /SV seed value constraints, when present.</summary>
    public PdfSignatureSeedValueInfo? SeedValue { get; }

    /// <summary>True when the owning signature field exposes a /Lock dictionary.</summary>
    public bool HasFieldLock => FieldLock is not null;

    /// <summary>True when the owning signature field exposes a /SV seed value dictionary.</summary>
    public bool HasSeedValue => SeedValue is not null;

    /// <summary>Signature handler /Filter name, for example Adobe.PPKLite, when readable.</summary>
    public string? Filter { get; }

    /// <summary>Signature /SubFilter name, for example adbe.pkcs7.detached, when readable.</summary>
    public string? SubFilter { get; }

    /// <summary>True when the /SubFilter is one of the common CMS, CAdES, or timestamp signature subfilters.</summary>
    public bool HasRecognizedSubFilter =>
        string.Equals(SubFilter, "adbe.pkcs7.detached", StringComparison.Ordinal) ||
        string.Equals(SubFilter, "adbe.pkcs7.sha1", StringComparison.Ordinal) ||
        string.Equals(SubFilter, "ETSI.CAdES.detached", StringComparison.Ordinal) ||
        string.Equals(SubFilter, "ETSI.RFC3161", StringComparison.Ordinal);

    /// <summary>True when the signature declares a detached CMS/PKCS#7 subfilter.</summary>
    public bool UsesDetachedCmsSubFilter =>
        string.Equals(SubFilter, "adbe.pkcs7.detached", StringComparison.Ordinal) ||
        string.Equals(SubFilter, "adbe.pkcs7.sha1", StringComparison.Ordinal);

    /// <summary>True when the signature declares an ETSI CAdES subfilter.</summary>
    public bool UsesCadesSubFilter => string.Equals(SubFilter, "ETSI.CAdES.detached", StringComparison.Ordinal);

    /// <summary>True when the signature declares an RFC 3161 document timestamp subfilter.</summary>
    public bool IsDocumentTimestamp => string.Equals(SubFilter, "ETSI.RFC3161", StringComparison.Ordinal);

    /// <summary>Signer /Name value, when present in the signature dictionary.</summary>
    public string? SignerName { get; }

    /// <summary>Signature /Location value, when present.</summary>
    public string? Location { get; }

    /// <summary>Signature /Reason value, when present.</summary>
    public string? Reason { get; }

    /// <summary>Signature /ContactInfo value, when present.</summary>
    public string? ContactInfo { get; }

    /// <summary>Raw signature /M signing time string, when present.</summary>
    public string? SigningTimeRaw { get; }

    /// <summary>True when the signature dictionary contains a readable /ByteRange array.</summary>
    public bool HasByteRange { get; }

    /// <summary>Exact numeric values read from the signature dictionary /ByteRange array.</summary>
    public IReadOnlyList<long> ByteRangeValues { get; }

    /// <summary>Number of numeric values found in the signature dictionary /ByteRange array.</summary>
    public int ByteRangeValueCount { get; }

    /// <summary>Number of byte ranges represented by the numeric /ByteRange values.</summary>
    public int ByteRangeSegmentCount => ByteRangeValueCount / 2;

    /// <summary>True when the signature dictionary contains a /Contents value.</summary>
    public bool HasContents { get; }

    /// <summary>Decoded `/Contents` bytes, including reserved padding, when readable.</summary>
    public byte[]? ContentsBytes { get; }

    /// <summary>Decoded /Contents byte count when the value could be read as a PDF string.</summary>
    public int? ContentsSizeBytes { get; }

    /// <summary>Encoded /Contents placeholder byte count, including delimiters, when readable from source syntax.</summary>
    public int? ContentsEncodedSizeBytes { get; }

    /// <summary>True when /Contents contains at least one decoded byte.</summary>
    public bool HasNonEmptyContents => ContentsSizeBytes.GetValueOrDefault() > 0;

    /// <summary>True when /Contents is a non-empty all-zero reservation ready for external signature bytes.</summary>
    public bool HasUnsignedContentsPlaceholder =>
        HasByteRange &&
        ContentsBytes is { Length: > 0 } &&
        ContentsBytes.All(static value => value == 0);

    /// <summary>Number of entries in the signature dictionary /Reference array, when readable.</summary>
    public int ReferenceCount { get; }
}
