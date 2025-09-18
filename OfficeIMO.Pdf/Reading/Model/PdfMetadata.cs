namespace OfficeIMO.Pdf;

/// <summary>
/// Basic PDF document metadata extracted from the Info dictionary.
/// </summary>
public sealed class PdfMetadata {
    /// <summary>Document title.</summary>
    public string? Title { get; init; }
    /// <summary>Document author.</summary>
    public string? Author { get; init; }
    /// <summary>Document subject.</summary>
    public string? Subject { get; init; }
    /// <summary>Document keywords.</summary>
    public string? Keywords { get; init; }
}
