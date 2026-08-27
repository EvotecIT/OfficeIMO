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
        double? layoutY = null,
        int? semanticNodeId = null,
        bool bidiVisualOrderResolved = false,
        OfficeTextDecorationStyle underlineStyle = OfficeTextDecorationStyle.None,
        OfficeTextDecorationStyle strikethroughStyle = OfficeTextDecorationStyle.None,
        OfficeTextBaseline baseline = OfficeTextBaseline.Normal)
        : this(text, x, y, width, height, font, color, alignment, lineHeight, paintOrder, linkUri, source, semanticRole, layoutY, semanticNodeId, null, bidiVisualOrderResolved, null, null, underlineStyle, strikethroughStyle, baseline) {
    }

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
        string? linkUri,
        string? source,
        string? semanticRole,
        double? layoutY,
        int? semanticNodeId,
        double? textAdvanceWidth,
        bool bidiVisualOrderResolved = false,
        int? semanticFragmentOrder = null,
        int? logicalTextOrder = null,
        OfficeTextDecorationStyle underlineStyle = OfficeTextDecorationStyle.None,
        OfficeTextDecorationStyle strikethroughStyle = OfficeTextDecorationStyle.None,
        OfficeTextBaseline baseline = OfficeTextBaseline.Normal)
        : base(HtmlRenderVisualKind.Text, x, y, width, height, paintOrder, linkUri, source, layoutY) {
        if (textAdvanceWidth.HasValue && (double.IsNaN(textAdvanceWidth.Value) || double.IsInfinity(textAdvanceWidth.Value))) {
            throw new ArgumentOutOfRangeException(nameof(textAdvanceWidth));
        }
        Text = text ?? throw new ArgumentNullException(nameof(text));
        Font = font;
        Color = color;
        Alignment = alignment;
        LineHeight = lineHeight;
        SemanticRole = semanticRole;
        SemanticNodeId = semanticNodeId;
        SemanticFragmentOrder = semanticFragmentOrder;
        LogicalTextOrder = logicalTextOrder;
        TextAdvanceWidth = textAdvanceWidth;
        BidiVisualOrderResolved = bidiVisualOrderResolved;
        UnderlineStyle = underlineStyle != OfficeTextDecorationStyle.None
            ? underlineStyle
            : font.IsUnderline ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None;
        StrikethroughStyle = strikethroughStyle != OfficeTextDecorationStyle.None
            ? strikethroughStyle
            : font.IsStrikethrough ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None;
        Baseline = baseline;
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

    /// <summary>Stable operation-scoped semantic node identifier shared by fragments from the same source element.</summary>
    public int? SemanticNodeId { get; }

    internal int? SemanticFragmentOrder { get; }

    internal int? LogicalTextOrder { get; }

    /// <summary>Resolved signed glyph advance for positioned inline text, distinct from its non-negative clipping frame.</summary>
    public double? TextAdvanceWidth { get; }

    /// <summary>Resolved CSS underline pattern.</summary>
    public OfficeTextDecorationStyle UnderlineStyle { get; }

    /// <summary>Resolved CSS strikethrough pattern.</summary>
    public OfficeTextDecorationStyle StrikethroughStyle { get; }

    /// <summary>Resolved CSS script baseline.</summary>
    public OfficeTextBaseline Baseline { get; }

    internal bool BidiVisualOrderResolved { get; }

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderText(Text, X + offsetX, Y + offsetY, Width, Height, Font, Color, Alignment, LineHeight, paintOrder, LinkUri, Source, SemanticRole, LayoutY + offsetY, SemanticNodeId, TextAdvanceWidth, BidiVisualOrderResolved, SemanticFragmentOrder, LogicalTextOrder, UnderlineStyle, StrikethroughStyle, Baseline);

    internal override HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderText(Text, X + offsetX, Y + offsetY, Width, Height, Font, Color, Alignment, LineHeight, paintOrder, LinkUri, Source, SemanticRole, LayoutY, SemanticNodeId, TextAdvanceWidth, BidiVisualOrderResolved, SemanticFragmentOrder, LogicalTextOrder, UnderlineStyle, StrikethroughStyle, Baseline);
}
