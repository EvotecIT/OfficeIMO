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
        bool bidiVisualOrderResolved = false)
        : this(text, x, y, width, height, font, color, alignment, lineHeight, paintOrder,
            linkUri, source, semanticRole, layoutY, semanticNodeId, null, bidiVisualOrderResolved, null, null,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, OfficeTextBaseline.Normal) {
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
        bool bidiVisualOrderResolved,
        int? semanticFragmentOrder,
        int? logicalTextOrder)
        : this(text, x, y, width, height, font, color, alignment, lineHeight, paintOrder,
            linkUri, source, semanticRole, layoutY, semanticNodeId, textAdvanceWidth, bidiVisualOrderResolved,
            semanticFragmentOrder, logicalTextOrder, OfficeTextDecorationStyle.None,
            OfficeTextDecorationStyle.None, OfficeTextBaseline.Normal) {
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
        OfficeTextBaseline baseline = OfficeTextBaseline.Normal,
        int baselineLevel = 0,
        double baselineScale = 1D,
        double baselineOffset = 0D,
        double? textPaintWidth = null,
        OfficeColor? decorationColor = null)
        : base(HtmlRenderVisualKind.Text, x, y, width, height, paintOrder, linkUri, source, layoutY) {
        if (textAdvanceWidth.HasValue && (double.IsNaN(textAdvanceWidth.Value) || double.IsInfinity(textAdvanceWidth.Value))) {
            throw new ArgumentOutOfRangeException(nameof(textAdvanceWidth));
        }
        if (textPaintWidth.HasValue &&
            (double.IsNaN(textPaintWidth.Value) || double.IsInfinity(textPaintWidth.Value) || textPaintWidth.Value < 0D)) {
            throw new ArgumentOutOfRangeException(nameof(textPaintWidth));
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
        TextPaintWidth = textPaintWidth;
        BidiVisualOrderResolved = bidiVisualOrderResolved;
        UnderlineStyle = underlineStyle != OfficeTextDecorationStyle.None
            ? underlineStyle
            : font.IsUnderline ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None;
        StrikethroughStyle = strikethroughStyle != OfficeTextDecorationStyle.None
            ? strikethroughStyle
            : font.IsStrikethrough ? OfficeTextDecorationStyle.Single : OfficeTextDecorationStyle.None;
        DecorationColor = decorationColor ?? color;
        Baseline = baseline;
        BaselineLevel = baselineLevel != 0
            ? baselineLevel
            : baseline == OfficeTextBaseline.Superscript ? 1
            : baseline == OfficeTextBaseline.Subscript ? -1 : 0;
        BaselineScale = baselineScale;
        BaselineOffset = baselineOffset;
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

    /// <summary>Measured glyph-paint width before CSS letter and word spacing are added to the positioned advance.</summary>
    public double? TextPaintWidth { get; }

    /// <summary>Resolved CSS underline pattern.</summary>
    public OfficeTextDecorationStyle UnderlineStyle { get; }

    /// <summary>Resolved CSS strikethrough pattern.</summary>
    public OfficeTextDecorationStyle StrikethroughStyle { get; }

    /// <summary>Resolved CSS color used to paint underlines and strikethroughs.</summary>
    public OfficeColor DecorationColor { get; }

    /// <summary>Resolved CSS script baseline.</summary>
    public OfficeTextBaseline Baseline { get; }

    /// <summary>Resolved cumulative CSS script nesting level.</summary>
    public int BaselineLevel { get; }

    /// <summary>Resolved cumulative CSS script font-size scale.</summary>
    public double BaselineScale { get; }

    /// <summary>Resolved cumulative CSS baseline displacement in layout units; negative values raise text.</summary>
    public double BaselineOffset { get; }

    internal bool BidiVisualOrderResolved { get; }

    internal override HtmlRenderVisual Translate(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderText(Text, X + offsetX, Y + offsetY, Width, Height, Font, Color, Alignment, LineHeight, paintOrder, LinkUri, Source, SemanticRole, LayoutY + offsetY, SemanticNodeId, TextAdvanceWidth, BidiVisualOrderResolved, SemanticFragmentOrder, LogicalTextOrder, UnderlineStyle, StrikethroughStyle, Baseline, BaselineLevel, BaselineScale, BaselineOffset, TextPaintWidth, DecorationColor);

    internal override HtmlRenderVisual TranslatePaint(double offsetX, double offsetY, int paintOrder) =>
        new HtmlRenderText(Text, X + offsetX, Y + offsetY, Width, Height, Font, Color, Alignment, LineHeight, paintOrder, LinkUri, Source, SemanticRole, LayoutY, SemanticNodeId, TextAdvanceWidth, BidiVisualOrderResolved, SemanticFragmentOrder, LogicalTextOrder, UnderlineStyle, StrikethroughStyle, Baseline, BaselineLevel, BaselineScale, BaselineOffset, TextPaintWidth, DecorationColor);
}
