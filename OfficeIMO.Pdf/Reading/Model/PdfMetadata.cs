namespace OfficeIMO.Pdf;

/// <summary>
/// Basic PDF document metadata extracted from the Info dictionary.
/// </summary>
public sealed class PdfMetadata {
    /// <summary>Document title.</summary>
    public string? Title { get; set; }
    /// <summary>Document author.</summary>
    public string? Author { get; set; }
    /// <summary>Document subject.</summary>
    public string? Subject { get; set; }
    /// <summary>Document keywords.</summary>
    public string? Keywords { get; set; }
    /// <summary>Print trapping status from the Info dictionary.</summary>
    public PdfTrappingStatus? TrappingStatus { get; set; }
    /// <summary>Document creation date from the Info dictionary.</summary>
    public DateTimeOffset? CreationDate { get; set; }
    /// <summary>Document modification date from the Info dictionary.</summary>
    public DateTimeOffset? ModificationDate { get; set; }
    /// <summary>PDF/X version from <c>GTS_PDFXVersion</c> in the Info dictionary.</summary>
    public string? PdfXVersion { get; set; }
    /// <summary>PDF/X conformance from <c>GTS_PDFXConformance</c> in the Info dictionary.</summary>
    public string? PdfXConformance { get; set; }
}
