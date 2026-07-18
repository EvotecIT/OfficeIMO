using OfficeIMO.Drawing;

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
    /// <summary>PDF base font name resolved from the active resource dictionary, when available.</summary>
    public string? BaseFont { get; }
    /// <summary>Font size in points.</summary>
    public double FontSize { get; }
    /// <summary>X position (points) in page user space.</summary>
    public double X { get; }
    /// <summary>Y position (points) in page user space.</summary>
    public double Y { get; }
    /// <summary>Advance width in user space for this span (includes spacing and hscale).</summary>
    public double Advance { get; }
    /// <summary>Fill color used for visual text rendering, when it could be read from the content stream.</summary>
    public OfficeColor? Color { get; }
    /// <summary>True when the text span should be painted during visual rendering.</summary>
    public bool IsVisible { get; }
    /// <summary>Baseline rotation in PDF user-space degrees, where positive angles are counter-clockwise.</summary>
    public double RotationDegrees { get; }
    /// <summary>Optional visual clipping path in page top-left coordinates.</summary>
    internal PdfPageClipPath? ClipPath { get; }
    internal string? DrawingFontFamily { get; }
    internal double PaintOrder { get; }
    internal int LogicalLineBreaksBefore { get; }
    internal bool LogicalLeadingSpace { get; }
    internal bool LogicalTrailingSpace { get; }
    /// <summary>Creates a new text span.</summary>
    public PdfTextSpan(string text, string fontResource, double fontSize, double x, double y, double advance = 0, OfficeColor? color = null, bool isVisible = true, double rotationDegrees = 0D, string? baseFont = null)
        : this(text, fontResource, fontSize, x, y, advance, color, isVisible, rotationDegrees, baseFont, null) {
    }

    internal PdfTextSpan(string text, string fontResource, double fontSize, double x, double y, double advance, OfficeColor? color, bool isVisible, double rotationDegrees, string? baseFont, PdfPageClipPath? clipPath, double paintOrder = 0D, string? drawingFontFamily = null, int logicalLineBreaksBefore = 0, bool logicalLeadingSpace = false, bool logicalTrailingSpace = false) {
        Text = text; FontResource = fontResource; BaseFont = baseFont; FontSize = fontSize; X = x; Y = y; Advance = advance; Color = color; IsVisible = isVisible; RotationDegrees = rotationDegrees; ClipPath = clipPath; PaintOrder = paintOrder; DrawingFontFamily = drawingFontFamily; LogicalLineBreaksBefore = logicalLineBreaksBefore; LogicalLeadingSpace = logicalLeadingSpace; LogicalTrailingSpace = logicalTrailingSpace;
    }
}
