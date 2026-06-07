namespace OfficeIMO.Pdf;

/// <summary>
/// Lightweight metadata from a PDF Document Security Store (/DSS) dictionary.
/// </summary>
public sealed class PdfDocumentDssInfo {
    internal PdfDocumentDssInfo(
        bool hasDss,
        int? objectNumber,
        IReadOnlyList<string> vriKeys,
        IReadOnlyList<int> certificateObjectNumbers,
        IReadOnlyList<int> ocspObjectNumbers,
        IReadOnlyList<int> crlObjectNumbers,
        IReadOnlyList<int> vriCertificateObjectNumbers,
        IReadOnlyList<int> vriOcspObjectNumbers,
        IReadOnlyList<int> vriCrlObjectNumbers,
        IReadOnlyList<int> timestampObjectNumbers) {
        HasDss = hasDss;
        ObjectNumber = objectNumber;
        VriKeys = vriKeys;
        CertificateObjectNumbers = certificateObjectNumbers;
        OcspObjectNumbers = ocspObjectNumbers;
        CrlObjectNumbers = crlObjectNumbers;
        VriCertificateObjectNumbers = vriCertificateObjectNumbers;
        VriOcspObjectNumbers = vriOcspObjectNumbers;
        VriCrlObjectNumbers = vriCrlObjectNumbers;
        TimestampObjectNumbers = timestampObjectNumbers;
    }

    internal static PdfDocumentDssInfo Empty { get; } = new PdfDocumentDssInfo(
        false,
        null,
        Array.Empty<string>(),
        Array.Empty<int>(),
        Array.Empty<int>(),
        Array.Empty<int>(),
        Array.Empty<int>(),
        Array.Empty<int>(),
        Array.Empty<int>(),
        Array.Empty<int>());

    /// <summary>True when the catalog exposes a /DSS dictionary.</summary>
    public bool HasDss { get; }

    /// <summary>Catalog /DSS object number, when the store is referenced indirectly.</summary>
    public int? ObjectNumber { get; }

    /// <summary>Keys present in the /VRI validation-related information dictionary.</summary>
    public IReadOnlyList<string> VriKeys { get; }

    /// <summary>Top-level /DSS /Certs object references, when readable.</summary>
    public IReadOnlyList<int> CertificateObjectNumbers { get; }

    /// <summary>Top-level /DSS /OCSPs object references, when readable.</summary>
    public IReadOnlyList<int> OcspObjectNumbers { get; }

    /// <summary>Top-level /DSS /CRLs object references, when readable.</summary>
    public IReadOnlyList<int> CrlObjectNumbers { get; }

    /// <summary>Certificate references discovered inside /VRI entries.</summary>
    public IReadOnlyList<int> VriCertificateObjectNumbers { get; }

    /// <summary>OCSP references discovered inside /VRI entries.</summary>
    public IReadOnlyList<int> VriOcspObjectNumbers { get; }

    /// <summary>CRL references discovered inside /VRI entries.</summary>
    public IReadOnlyList<int> VriCrlObjectNumbers { get; }

    /// <summary>Timestamp references discovered inside /VRI entries.</summary>
    public IReadOnlyList<int> TimestampObjectNumbers { get; }

    /// <summary>Number of /VRI validation-related information entries.</summary>
    public int VriEntryCount => VriKeys.Count;

    /// <summary>Total top-level DSS evidence object references.</summary>
    public int TopLevelEvidenceObjectCount => CertificateObjectNumbers.Count + OcspObjectNumbers.Count + CrlObjectNumbers.Count;

    /// <summary>Total VRI evidence object references.</summary>
    public int VriEvidenceObjectCount => VriCertificateObjectNumbers.Count + VriOcspObjectNumbers.Count + VriCrlObjectNumbers.Count + TimestampObjectNumbers.Count;

    /// <summary>True when the DSS exposes at least one evidence object reference or VRI entry.</summary>
    public bool HasValidationEvidence => HasDss && (TopLevelEvidenceObjectCount > 0 || VriEntryCount > 0 || VriEvidenceObjectCount > 0);
}
