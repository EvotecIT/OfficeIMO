namespace OfficeIMO.Pdf;

/// <summary>A bounded, reviewable item discovered by the PDF sanitization engine.</summary>
public sealed class PdfSanitizationFinding {
    internal PdfSanitizationFinding(PdfSanitizationFindingKind kind, int objectNumber, string path, string detail, PdfSanitizationActionKind? actionKind = null) {
        Kind = kind;
        ObjectNumber = objectNumber;
        Path = path;
        Detail = detail;
        ActionKind = actionKind;
    }

    /// <summary>Finding category.</summary>
    public PdfSanitizationFindingKind Kind { get; }

    /// <summary>Source indirect object number, or zero for a direct trailer-owned object.</summary>
    public int ObjectNumber { get; }

    /// <summary>Stable object-relative path to the unsafe entry.</summary>
    public string Path { get; }

    /// <summary>Action type, URI, or payload marker that caused the finding.</summary>
    public string Detail { get; }

    /// <summary>Typed action category when this finding represents an action or URI target.</summary>
    public PdfSanitizationActionKind? ActionKind { get; }
}
