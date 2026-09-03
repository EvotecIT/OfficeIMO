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
    /// <summary>Active content-stream-scoped marked-content identifier from a tagged PDF, when present.</summary>
    public int? MarkedContentId { get; }
    /// <summary>Form XObject or other content-stream object owning the MCID, or null for page content.</summary>
    public int? ContentStreamObjectNumber { get; }
    /// <summary>Optional visual clipping path in page top-left coordinates.</summary>
    internal PdfPageClipPath? ClipPath { get; }
    internal string? DrawingFontFamily { get; }
    internal double PaintOrder { get; }
    internal PdfContentOrderKey? ContentOrderKey { get; }
    internal PdfContentOrderKey? TextObjectOrderKey { get; }
    internal int LogicalLineBreaksBefore { get; }
    internal bool LogicalLeadingSpace { get; }
    internal bool LogicalTrailingSpace { get; }
    internal IReadOnlyList<double>? CharacterAdvances { get; }
    internal IReadOnlyList<int>? GlyphCharacterLengths { get; }
    internal int TextRenderingMode { get; }
    internal bool CanRestamp { get; }
    internal bool CanScaleAggregateAdvance { get; }
    internal double RestampFontSize { get; }
    internal string RestampText { get; }
    internal Matrix2D? TextToPageTransform { get; }
    internal string? VisualPaintIdentity { get; }
    /// <summary>Creates a new text span.</summary>
    public PdfTextSpan(string text, string fontResource, double fontSize, double x, double y, double advance = 0, OfficeColor? color = null, bool isVisible = true, double rotationDegrees = 0D, string? baseFont = null)
        : this(text, fontResource, fontSize, x, y, advance, color, isVisible, rotationDegrees, baseFont, null) {
    }

    internal PdfTextSpan(string text, string fontResource, double fontSize, double x, double y, double advance, OfficeColor? color, bool isVisible, double rotationDegrees, string? baseFont, PdfPageClipPath? clipPath, double paintOrder = 0D, string? drawingFontFamily = null, int logicalLineBreaksBefore = 0, bool logicalLeadingSpace = false, bool logicalTrailingSpace = false, PdfContentOrderKey? contentOrderKey = null, IReadOnlyList<double>? characterAdvances = null, int textRenderingMode = 0, bool canRestamp = true, double? restampFontSize = null, string? restampText = null, bool canScaleAggregateAdvance = true, int? markedContentId = null, int? contentStreamObjectNumber = null, PdfContentOrderKey? textObjectOrderKey = null, Matrix2D? textToPageTransform = null, string? visualPaintIdentity = null, IReadOnlyList<int>? glyphCharacterLengths = null) {
        Text = text; FontResource = fontResource; BaseFont = baseFont; FontSize = fontSize; X = x; Y = y; Advance = advance; Color = color; IsVisible = isVisible; RotationDegrees = rotationDegrees; MarkedContentId = markedContentId; ContentStreamObjectNumber = contentStreamObjectNumber; ClipPath = clipPath; PaintOrder = paintOrder; DrawingFontFamily = drawingFontFamily; LogicalLineBreaksBefore = logicalLineBreaksBefore; LogicalLeadingSpace = logicalLeadingSpace; LogicalTrailingSpace = logicalTrailingSpace; ContentOrderKey = contentOrderKey; TextObjectOrderKey = textObjectOrderKey; CharacterAdvances = characterAdvances?.ToArray(); GlyphCharacterLengths = glyphCharacterLengths?.ToArray(); TextRenderingMode = textRenderingMode; CanRestamp = canRestamp; CanScaleAggregateAdvance = canScaleAggregateAdvance; RestampFontSize = restampFontSize ?? fontSize; RestampText = restampText ?? text; TextToPageTransform = textToPageTransform; VisualPaintIdentity = visualPaintIdentity;
    }

    internal PdfTextSpan WithCanRestamp(bool canRestamp) => new PdfTextSpan(
        Text, FontResource, FontSize, X, Y, Advance, Color, IsVisible, RotationDegrees, BaseFont, ClipPath,
        PaintOrder, DrawingFontFamily, LogicalLineBreaksBefore, LogicalLeadingSpace, LogicalTrailingSpace,
        ContentOrderKey, CharacterAdvances, TextRenderingMode, canRestamp, RestampFontSize, RestampText, CanScaleAggregateAdvance, MarkedContentId, ContentStreamObjectNumber, TextObjectOrderKey, TextToPageTransform, VisualPaintIdentity, GlyphCharacterLengths);

    internal PdfTextSpan WithOffset(double deltaX, double deltaY) => new PdfTextSpan(
        Text, FontResource, FontSize, X + deltaX, Y + deltaY, Advance, Color, IsVisible, RotationDegrees, BaseFont, ClipPath,
        PaintOrder, DrawingFontFamily, LogicalLineBreaksBefore, LogicalLeadingSpace, LogicalTrailingSpace,
        ContentOrderKey, CharacterAdvances, TextRenderingMode, CanRestamp, RestampFontSize, RestampText, CanScaleAggregateAdvance, MarkedContentId, ContentStreamObjectNumber, TextObjectOrderKey, TextToPageTransform, VisualPaintIdentity, GlyphCharacterLengths);

    internal PdfTextSpan WithVisualFontSize(double fontSize) => new PdfTextSpan(
        Text, FontResource, fontSize, X, Y, Advance, Color, IsVisible, RotationDegrees, BaseFont, ClipPath,
        PaintOrder, DrawingFontFamily, LogicalLineBreaksBefore, LogicalLeadingSpace, LogicalTrailingSpace,
        ContentOrderKey, CharacterAdvances, TextRenderingMode, CanRestamp, RestampFontSize, RestampText, CanScaleAggregateAdvance, MarkedContentId, ContentStreamObjectNumber, TextObjectOrderKey, TextToPageTransform, VisualPaintIdentity, GlyphCharacterLengths);

    internal bool CanProjectCompleteText(double? pageHeight) {
        if (!IsVisible || string.IsNullOrEmpty(Text)) return false;
        if (!ClipPath.HasValue) return true;
        if (!pageHeight.HasValue || Math.Abs(RotationDegrees) > 0.01D) return false;

        PdfPageClipPath clip = ClipPath.Value;
        if (!clip.IsRectangle || !clip.IsExact || clip.ContainsTextClipping) return false;
        // Extracted spans expose a baseline rather than font ascent/descent metrics. Use the painted
        // glyph box approximation here so tight producer cell clips do not look like partial text.
        const double approximateAscentFactor = 0.8D;
        double width = Advance > 0D ? Advance : Text.Length * FontSize * 0.55D;
        double height = Math.Max(1D, FontSize);
        double left = X;
        double top = pageHeight.Value - Y - FontSize * approximateAscentFactor;
        const double tolerance = 0.05D;
        return left + tolerance >= clip.X &&
               top + tolerance >= clip.Y &&
               left + width <= clip.X + clip.Width + tolerance &&
               top + height <= clip.Y + clip.Height + tolerance;
    }
}
