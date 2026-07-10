using OfficeIMO.Drawing;

namespace OfficeIMO.Html;

/// <summary>
/// Positioned text segment retained as text for image and PDF backends.
/// </summary>
public sealed class HtmlRenderText : HtmlRenderVisual {
    internal HtmlRenderText(
        string text,
        double x,
        double y,
        double width,
        double height,
        OfficeFontInfo font,
        OfficeColor color,
        OfficeTextAlignment alignment,
        double lineHeight,
        int paintOrder,
        string? linkUri = null,
        string? source = null,
        string? semanticRole = null,
        double? layoutY = null)
        : base(HtmlRenderVisualKind.Text, x, y, width, height, paintOrder, linkUri, source, layoutY) {
        Text = text ?? throw new ArgumentNullException(nameof(text));
        Font = font;
        Color = color;
        Alignment = alignment;
        LineHeight = lineHeight;
        SemanticRole = semanticRole;
    }

    /// <summary>Text content represented by this visual segment.</summary>
    public string Text { get; }

    /// <summary>Resolved font descriptor.</summary>
    public OfficeFontInfo Font { get; }

    /// <summary>Resolved text color.</summary>
    public OfficeColor Color { get; }

    /// <summary>Resolved horizontal text alignment.</summary>
    public OfficeTextAlignment Alignment { get; }

    /// <summary>Resolved line height in CSS pixels.</summary>
    public double LineHeight { get; }

    /// <summary>Optional semantic role such as heading, paragraph, or list item.</summary>
    public string? SemanticRole { get; }

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderText(Text, X + offsetX, Y + offsetY, Width, Height, Font, Color, Alignment, LineHeight, paintOrder, LinkUri, Source, SemanticRole, LayoutY + offsetY);

    internal override HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderText(Text, X + offsetX, Y + offsetY, Width, Height, Font, Color, Alignment, LineHeight, paintOrder, LinkUri, Source, SemanticRole, LayoutY);
}
