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
}
