namespace OfficeIMO.Pdf;

/// <summary>
/// PDF/X identification metadata written to the <c>pdfxid</c> XMP namespace.
/// </summary>
public sealed class PdfXIdentification {
    /// <summary>PDF/X identification XMP namespace.</summary>
    public const string NamespaceUri = "http://www.npes.org/pdfx/ns/id/";

    /// <summary>Creates PDF/X identification metadata.</summary>
    public PdfXIdentification(string version, string? conformance = null) {
        Guard.NotNullOrWhiteSpace(version, nameof(version));
        if (conformance != null && conformance.Length == 0) {
            throw new ArgumentException("PDF/X conformance cannot be empty.", nameof(conformance));
        }

        Version = version.Trim();
        Conformance = conformance?.Trim();
    }

    /// <summary>PDF/X version, for example <c>PDF/X-4</c>.</summary>
    public string Version { get; }

    /// <summary>Optional PDF/X conformance identifier required by older PDF/X profiles.</summary>
    public string? Conformance { get; }

    /// <summary>Creates PDF/X-1a:2003 identification metadata.</summary>
    public static PdfXIdentification PdfX1A2003() =>
        new("PDF/X-1a:2003", "PDF/X-1a:2003");

    /// <summary>Creates PDF/X-4 identification metadata.</summary>
    public static PdfXIdentification PdfX4() => new("PDF/X-4");

    internal PdfXIdentification Clone() => new(Version, Conformance);
}
