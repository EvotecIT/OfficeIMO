namespace OfficeIMO.Pdf;

/// <summary>
/// A piece of text extracted from a PDF page with basic font and position info.
/// Coordinates are in user space units (points) as emitted by content stream Tm/Td.
/// </summary>
public sealed class PdfTextSpan {
    /// <summary>Text content of the span.</summary>
    public string Text { get; }
    /// <summary>Font resource name from the page resources (e.g., F1).</summary>
    public string FontResource { get; }
    /// <summary>Font size in points.</summary>
    public double FontSize { get; }
    /// <summary>X position (points) in page user space.</summary>
    public double X { get; }
    /// <summary>Y position (points) in page user space.</summary>
    public double Y { get; }
    /// <summary>Advance width in user space for this span (includes spacing and hscale).</summary>
    public double Advance { get; }
    /// <summary>Creates a new text span.</summary>
    public PdfTextSpan(string text, string fontResource, double fontSize, double x, double y, double advance = 0) {
        Text = text; FontResource = fontResource; FontSize = fontSize; X = x; Y = y; Advance = advance;
    }
}
