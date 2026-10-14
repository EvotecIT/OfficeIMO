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
        OfficeColor? decorationColor = null,
        OfficeTextFeatureSettings? featureSettings = null,
        string? fontPalette = null,
        double? layoutHeight = null,
        OfficeFontFaceDescriptor? fontDescriptor = null,
        double paintTopOverflow = 0D,
        bool paintOnly = false,
        HtmlRenderRectangle? linkBounds = null)
        : base(HtmlRenderVisualKind.Text, x, y, width, height, paintOrder, linkUri, source, layoutY, layoutHeight) {
        if (textAdvanceWidth.HasValue && (double.IsNaN(textAdvanceWidth.Value) || double.IsInfinity(textAdvanceWidth.Value))) {
            throw new ArgumentOutOfRangeException(nameof(textAdvanceWidth));
        }
        if (textPaintWidth.HasValue &&
            (double.IsNaN(textPaintWidth.Value) || double.IsInfinity(textPaintWidth.Value) || textPaintWidth.Value < 0D)) {
            throw new ArgumentOutOfRangeException(nameof(textPaintWidth));
        }
        if (paintTopOverflow < 0D || double.IsNaN(paintTopOverflow) || double.IsInfinity(paintTopOverflow)) {
            throw new ArgumentOutOfRangeException(nameof(paintTopOverflow));
        }
        Text = text ?? throw new ArgumentNullException(nameof(text));
        IsPaintOnly = paintOnly;
        LinkBounds = linkBounds;
        Font = font;
        FontDescriptor = fontDescriptor ?? OfficeFontFaceDescriptor.FromStyle(font.Style);
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
        FeatureSettings = featureSettings ?? OfficeTextFeatureSettings.Default;
        FontPalette = string.IsNullOrWhiteSpace(fontPalette) ? "normal" : fontPalette!;
        Baseline = baseline;
        BaselineLevel = baselineLevel != 0
            ? baselineLevel
            : baseline == OfficeTextBaseline.Superscript ? 1
            : baseline == OfficeTextBaseline.Subscript ? -1 : 0;
        BaselineScale = baselineScale;
        BaselineOffset = baselineOffset;
        PaintTopOverflow = paintTopOverflow;
    }

    /// <summary>Text content represented by this visual segment.</summary>
    public string Text { get; }

    /// <summary>Resolved font family, size, and compatibility style flags.</summary>
    public OfficeFontInfo Font { get; }

    /// <summary>Resolved CSS face weight, stretch, and slant requested for this text.</summary>
    public OfficeFontFaceDescriptor FontDescriptor { get; }

    /// <summary>Resolved text color.</summary>
    public OfficeColor Color { get; }

    /// <summary>Resolved horizontal text alignment.</summary>
    public OfficeTextAlignment Alignment { get; }

    /// <summary>Resolved line height in CSS pixels.</summary>
    public double LineHeight { get; }

    // Positioned paint adapters require a positive carrier even when CSS keeps
    // a legal zero line metric. The adapter never uses this to advance HTML flow.
    internal double PaintLineHeight => LineHeight == 0D ? 0.01D : LineHeight;

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

    /// <summary>Resolved OpenType feature selections for this text run.</summary>
    public OfficeTextFeatureSettings FeatureSettings { get; }

    /// <summary>Resolved CSS font-palette value.</summary>
    public string FontPalette { get; }

    /// <summary>Resolved CSS script baseline.</summary>
    public OfficeTextBaseline Baseline { get; }

    /// <summary>Resolved cumulative CSS script nesting level.</summary>
    public int BaselineLevel { get; }

    /// <summary>Resolved cumulative CSS script font-size scale.</summary>
    public double BaselineScale { get; }

    /// <summary>Resolved cumulative CSS baseline displacement in layout units; negative values raise text.</summary>
    public double BaselineOffset { get; }

    // Selected face ascent can extend above the one-em baseline anchor used by positioned text.
    internal double PaintTopOverflow { get; }

    // Linked glyph paint and advance exclude paint-frame tolerance and extra line leading.
    internal HtmlRenderRectangle? LinkBounds { get; }

    private HtmlRenderRectangle? TranslateLinkBounds(double x, double y) =>
        LinkBounds is HtmlRenderRectangle bounds
            ? new HtmlRenderRectangle(bounds.X + x, bounds.Y + y, bounds.Width, bounds.Height)
            : null;

    /// <summary>Shadow clones paint text but do not contribute to scrollable layout overflow.</summary>
    internal bool IsPaintOnly { get; }

    internal HtmlRenderText ResolveBaselineForPainting() =>
        Baseline == OfficeTextBaseline.Normal && BaselineScale == 1D && BaselineOffset == 0D
            ? this
            : new HtmlRenderText(Text, X, Y + BaselineOffset, Width, Height,
                Font.WithSize(Font.Size * BaselineScale), Color, Alignment, LineHeight, PaintOrder,
                LinkUri, Source, SemanticRole, LayoutY, SemanticNodeId, TextAdvanceWidth,
                BidiVisualOrderResolved, SemanticFragmentOrder, LogicalTextOrder,
                UnderlineStyle, StrikethroughStyle, OfficeTextBaseline.Normal, 0, 1D, 0D,
                TextPaintWidth, DecorationColor, FeatureSettings, FontPalette, LayoutHeight, FontDescriptor,
                PaintTopOverflow, IsPaintOnly, LinkBounds).WithWrappedLines(WrappedLines);

    // Generated margin frames retain their complete text in the scene. Both
    // painting adapters consume the same measured lines instead of rewrapping.
    internal IReadOnlyList<OfficeTextLine>? WrappedLines { get; }

    private HtmlRenderText(HtmlRenderText source, IReadOnlyList<OfficeTextLine> wrappedLines)
        : this(source.Text, source.X, source.Y, source.Width, source.Height, source.Font,
            source.Color, source.Alignment, source.LineHeight, source.PaintOrder, source.LinkUri,
            source.Source, source.SemanticRole, source.LayoutY, source.SemanticNodeId,
            source.TextAdvanceWidth, source.BidiVisualOrderResolved, source.SemanticFragmentOrder,
            source.LogicalTextOrder, source.UnderlineStyle, source.StrikethroughStyle, source.Baseline,
            source.BaselineLevel, source.BaselineScale, source.BaselineOffset, source.TextPaintWidth,
            source.DecorationColor, source.FeatureSettings, source.FontPalette, source.LayoutHeight,
            source.FontDescriptor, source.PaintTopOverflow, source.IsPaintOnly, source.LinkBounds) {
        WrappedLines = wrappedLines;
    }

    internal HtmlRenderText WithWrappedLines(IReadOnlyList<OfficeTextLine>? lines) =>
        lines == null ? this : new HtmlRenderText(this, lines);

    internal IEnumerable<HtmlRenderText> GetWrappedPaintFragments() {
        if (WrappedLines == null) {
            yield return this;
            yield break;
        }
        for (int index = 0; index < WrappedLines.Count; index++) {
            double offsetY = index * LineHeight;
            if (offsetY >= Height) yield break;
            OfficeTextLine line = WrappedLines[index];
            if (line.Text.Length == 0) continue;
            yield return new HtmlRenderText(line.Text, X + line.OffsetX, Y + offsetY,
                Math.Max(0.01D, Width - line.OffsetX), Math.Min(LineHeight, Height - offsetY),
                Font, Color, Alignment, LineHeight, PaintOrder, LinkUri, Source, SemanticRole,
                LayoutY + offsetY, SemanticNodeId, line.Width, BidiVisualOrderResolved,
                SemanticFragmentOrder, LogicalTextOrder, UnderlineStyle, StrikethroughStyle,
                Baseline, BaselineLevel, BaselineScale, BaselineOffset, null,
                DecorationColor, FeatureSettings, FontPalette, Math.Min(LineHeight, LayoutHeight),
                FontDescriptor, PaintTopOverflow, IsPaintOnly, LinkBounds);
        }
    }

    internal bool BidiVisualOrderResolved { get; }

    internal override HtmlRenderVisual TranslateCore(double offsetX, double offsetY, int paintOrder) =>
        offsetX == 0D && offsetY == 0D && paintOrder == PaintOrder ? this :
        new HtmlRenderText(Text, X + offsetX, Y + offsetY, Width, Height, Font, Color, Alignment, LineHeight, paintOrder, LinkUri, Source, SemanticRole, LayoutY + offsetY, SemanticNodeId, TextAdvanceWidth, BidiVisualOrderResolved, SemanticFragmentOrder, LogicalTextOrder, UnderlineStyle, StrikethroughStyle, Baseline, BaselineLevel, BaselineScale, BaselineOffset, TextPaintWidth, DecorationColor, FeatureSettings, FontPalette, LayoutHeight, FontDescriptor, PaintTopOverflow, IsPaintOnly, TranslateLinkBounds(offsetX, offsetY)).WithWrappedLines(WrappedLines);

    internal override HtmlRenderVisual TranslatePaintCore(double offsetX, double offsetY, int paintOrder) =>
        offsetX == 0D && offsetY == 0D && paintOrder == PaintOrder ? this :
        new HtmlRenderText(Text, X + offsetX, Y + offsetY, Width, Height, Font, Color, Alignment, LineHeight, paintOrder, LinkUri, Source, SemanticRole, LayoutY, SemanticNodeId, TextAdvanceWidth, BidiVisualOrderResolved, SemanticFragmentOrder, LogicalTextOrder, UnderlineStyle, StrikethroughStyle, Baseline, BaselineLevel, BaselineScale, BaselineOffset, TextPaintWidth, DecorationColor, FeatureSettings, FontPalette, LayoutHeight, FontDescriptor, PaintTopOverflow, IsPaintOnly, TranslateLinkBounds(offsetX, offsetY)).WithWrappedLines(WrappedLines);
}
