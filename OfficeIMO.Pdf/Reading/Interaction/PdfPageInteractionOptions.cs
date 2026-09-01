namespace OfficeIMO.Pdf;

/// <summary>Bounds and inclusion policy for page interaction-map creation.</summary>
public sealed class PdfPageInteractionOptions {
    /// <summary>Maximum number of text-element regions emitted for one page.</summary>
    public int MaxTextRegions { get; set; } = 100000;

    /// <summary>Maximum number of image-placement regions emitted for one page.</summary>
    public int MaxImageRegions { get; set; } = 10000;

    /// <summary>Include invisible extracted text spans. Disabled by default.</summary>
    public bool IncludeInvisibleText { get; set; }
}
