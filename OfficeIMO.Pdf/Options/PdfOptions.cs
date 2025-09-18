namespace OfficeIMO.Pdf;

/// <summary>
/// Options controlling page geometry and default typography for a PDF document.
/// </summary>
public sealed class PdfOptions {
    /// <summary>Page width in points (1 pt = 1/72 in). Default is 612 (Letter 8.5in).</summary>
    public double PageWidth { get; set; } = 612; // Letter 8.5in * 72
    /// <summary>Page height in points. Default is 792 (Letter 11in).</summary>
    public double PageHeight { get; set; } = 792; // Letter 11in * 72
    /// <summary>Left margin in points. Default 72 (1 inch).</summary>
    public double MarginLeft { get; set; } = 72; // 1 in
    /// <summary>Right margin in points. Default 72 (1 inch).</summary>
    public double MarginRight { get; set; } = 72;
    /// <summary>Top margin in points. Default 72 (1 inch).</summary>
    public double MarginTop { get; set; } = 72;
    /// <summary>Bottom margin in points. Default 72 (1 inch).</summary>
    public double MarginBottom { get; set; } = 72;
    /// <summary>Default standard font used for paragraphs.</summary>
    public PdfStandardFont DefaultFont { get; set; } = PdfStandardFont.Courier;
    /// <summary>Default paragraph font size in points. Default 11.</summary>
    public double DefaultFontSize { get; set; } = 11;
}
