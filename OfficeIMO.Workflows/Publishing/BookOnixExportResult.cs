using OfficeIMO.Epub;
using System.Security.Cryptography;

namespace OfficeIMO.Workflows;

/// <summary>A schema-validated bibliographic ONIX record and the exact EPUB it describes.</summary>
public sealed class BookOnixExportResult {
    internal BookOnixExportResult(byte[] bytes, EpubWriteResult publication,
        IReadOnlyList<OfficeConversionFidelityDiagnostic> diagnostics, bool acknowledged) {
        Bytes = bytes; Publication = publication;
        PublicationSha256 = Convert.ToHexString(SHA256.HashData(publication.Bytes));
        ImportDiagnostics = Array.AsReadOnly(diagnostics.ToArray()); ImportLossAcknowledged = acknowledged;
    }
    /// <summary>UTF-8 ONIX 3.1 reference-tag XML, at most 1 MiB.</summary>
    public byte[] Bytes { get; }
    /// <summary>Exact EPUB bytes and canonical writer preservation/loss report.</summary>
    public EpubWriteResult Publication { get; }
    /// <summary>Uppercase SHA-256 of Publication.Bytes at export time; not a signature or retailer receipt.</summary>
    public string PublicationSha256 { get; }
    /// <summary>Retained import findings; schema validation does not erase conversion losses.</summary>
    public IReadOnlyList<OfficeConversionFidelityDiagnostic> ImportDiagnostics { get; }
    /// <summary>Whether non-fatal manuscript import losses were acknowledged.</summary>
    public bool ImportLossAcknowledged { get; }
}
