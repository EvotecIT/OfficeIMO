using System.Collections.Generic;
using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeDrawing {
    internal OfficeDrawing AddPositionedTextWithNaturalAdvance(
        string text, double x, double y, double width, double height,
        OfficeFontInfo? font, OfficeColor? color, double? lineHeight,
        OfficeTextDirection textDirection = OfficeTextDirection.Auto) =>
        AddPositionedTextWithResolvedDirection(text, x, y, width, height, font, color, lineHeight,
            textAdvanceWidth: null, textDirection: textDirection);

    internal OfficeDrawing AddPositionedTextWithResolvedDirection(
        string text, double x, double y, double width, double height,
        OfficeFontInfo? font, OfficeColor? color, double? lineHeight,
        double? textAdvanceWidth, OfficeTextDirection textDirection) =>
        AddTextCore(text, x, y, width, height, font, color, OfficeTextAlignment.Left, lineHeight,
            OfficeTextVerticalAlignment.Top, 0D, null, null, false, false, false, false, false, null, null,
            OfficeTextOverflowBehavior.Clip, textAdvanceWidth,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, OfficeTextBaseline.Normal,
            baselineLevel: 0, baselineScale: 1D, baselineOffset: 0D, decorationColor: null,
            featureSettings: null, fontPalette: null, allowOverflow: true, textDirection: textDirection,
            clipToFrame: false);

    internal OfficeDrawing AddClippedPositionedTextWithNaturalAdvance(
        string text, double x, double y, double width, double height,
        double clipX, double clipY, OfficeClipPath clipPath,
        OfficeFontInfo? font, OfficeColor? color, double? lineHeight,
        OfficeTextDirection textDirection = OfficeTextDirection.Auto) =>
        AddClippedPositionedTextWithResolvedDirection(text, x, y, width, height,
            clipX, clipY, clipPath, font, color, lineHeight,
            textAdvanceWidth: null, textDirection: textDirection);

    internal OfficeDrawing AddClippedPositionedTextWithResolvedDirection(
        string text, double x, double y, double width, double height,
        double clipX, double clipY, OfficeClipPath clipPath,
        OfficeFontInfo? font, OfficeColor? color, double? lineHeight,
        double? textAdvanceWidth, OfficeTextDirection textDirection) {
        if (clipPath == null) throw new ArgumentNullException(nameof(clipPath));
        ValidateFiniteNonNegative(clipX, nameof(clipX));
        ValidateFiniteNonNegative(clipY, nameof(clipY));
        if (clipX + clipPath.Width > Width || clipY + clipPath.Height > Height)
            throw new ArgumentOutOfRangeException(nameof(clipPath), "Text clip must fit inside the drawing bounds.");
        var clipped = new OfficeDrawing(Math.Max(0.01D, clipPath.Width), Math.Max(0.01D, clipPath.Height));
        clipped.AddTextCore(text, x - clipX, y - clipY, width, height, font, color, OfficeTextAlignment.Left, lineHeight,
            OfficeTextVerticalAlignment.Top, 0D, null, null, false, false, false, false, false, null, null,
            OfficeTextOverflowBehavior.Clip, textAdvanceWidth,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, OfficeTextBaseline.Normal,
            baselineLevel: 0, baselineScale: 1D, baselineOffset: 0D, decorationColor: null,
            featureSettings: null, fontPalette: null, allowOverflow: true, textDirection: textDirection,
            clipToFrame: false);
        return AddClippedDrawing(clipped, clipX, clipY, clipPath);
    }

    /// <summary>Marks the text run added last, directly or inside its clip group, as painted glyphs.</summary>
    internal void MarkLastTextAsPaintedGlyphs() {
        if (Elements.Count == 0) return;
        OfficeDrawingElement last = Elements[Elements.Count - 1];
        // Group.Drawing returns a copy; mark the retained inner drawing that renderers read.
        if (last is OfficeDrawingGroup group && group.InnerDrawing.Elements.Count > 0) {
            last = group.InnerDrawing.Elements[group.InnerDrawing.Elements.Count - 1];
        }
        if (last is OfficeDrawingText text) text.PreservesPaintedGlyphs = true;
    }

    /// <summary>Attaches retained glyphs to the last ordinary positioned text element.</summary>
    internal void MarkLastTextWithPreparedShaping(OfficePreparedTextShaping prepared) {
        if (prepared == null) throw new ArgumentNullException(nameof(prepared));
        if (Elements.Count == 0) throw new InvalidOperationException("No text element is available for prepared shaping.");
        OfficeDrawingElement last = Elements[Elements.Count - 1];
        if (last is OfficeDrawingGroup group && group.InnerDrawing.Elements.Count > 0)
            last = group.InnerDrawing.Elements[group.InnerDrawing.Elements.Count - 1];
        if (last is not OfficeDrawingText text || text.WrapText || text.ShrinkToFit || text.StackedText ||
            text.TextDirection == OfficeTextDirection.TopToBottom || text.RasterText.IndexOfAny(new[] { '\r', '\n' }) >= 0)
            throw new InvalidOperationException("Prepared shaping requires one already-positioned horizontal text slice.");
        if (!string.Equals(OfficeArabicTextShaper.ToLogicalText(text.Text), prepared.Text, StringComparison.Ordinal))
            throw new ArgumentException("Prepared shaping must describe the complete text element.", nameof(prepared));
        text.SetPaintedText(prepared.Text, prepared.Text);
        text.PreparedShaping = prepared;
    }

    internal void SetLastTextOutlineStroke(OfficeColor color, double width, bool fillSuppressed, IReadOnlyList<double>? dashPattern = null, double dashOffset = 0D,
        OfficeStrokeLineCap cap = OfficeStrokeLineCap.Butt, OfficeStrokeLineJoin join = OfficeStrokeLineJoin.Miter, double miterLimit = 10D) {
        if (Elements.Count == 0) return;
        OfficeDrawingElement last = Elements[Elements.Count - 1];
        if (last is OfficeDrawingGroup group && group.InnerDrawing.Elements.Count > 0)
            last = group.InnerDrawing.Elements[group.InnerDrawing.Elements.Count - 1];
        if (last is not OfficeDrawingText text) return;
        text.OutlineStrokeColor = color;
        text.OutlineStrokeWidth = width;
        text.FillSuppressed = fillSuppressed;
        text.OutlineStrokeDashPattern = dashPattern == null ? null : new List<double>(dashPattern).AsReadOnly();
        text.OutlineStrokeDashOffset = dashOffset;
        text.OutlineStrokeLineCap = cap;
        text.OutlineStrokeLineJoin = join;
        text.OutlineStrokeMiterLimit = miterLimit;
    }

    internal void SetLastTextPaintedProjection(string logicalText, string paintedText) {
        if (Elements.Count == 0) return;
        OfficeDrawingElement last = Elements[Elements.Count - 1];
        if (last is OfficeDrawingGroup group && group.InnerDrawing.Elements.Count > 0)
            last = group.InnerDrawing.Elements[group.InnerDrawing.Elements.Count - 1];
        if (last is OfficeDrawingText text) text.SetPaintedText(logicalText, paintedText);
    }

    /// <summary>Retains an affine glyph placement without transforming its outline pen.</summary>
    internal void SetLastTextGlyphTransform(OfficeTransform transform) {
        if (Elements.Count == 0) return;
        OfficeDrawingElement last = Elements[Elements.Count - 1];
        if (last is OfficeDrawingGroup group && group.InnerDrawing.Elements.Count > 0)
            last = group.InnerDrawing.Elements[group.InnerDrawing.Elements.Count - 1];
        if (last is OfficeDrawingText text) text.GlyphTransform = transform;
    }

    /// <summary>Adds a resolved source text run with its glyph advance and a destination-space rotation or mirror transform. Source text is clipped, never shortened with an ellipsis.</summary>
    public OfficeDrawing AddPositionedText(
        string text, double x, double y, double width, double height,
        OfficeImageFrameTransform frameTransform,
        OfficeFontInfo? font = null, OfficeColor? color = null,
        OfficeTextAlignment alignment = OfficeTextAlignment.Left,
        double? lineHeight = null, double? textAdvanceWidth = null) =>
        AddTextCore(text, x, y, width, height, font, color, alignment, lineHeight,
            OfficeTextVerticalAlignment.Top, frameTransform.RotationDegrees, frameTransform.CenterX, frameTransform.CenterY,
            false, false, false, frameTransform.FlipHorizontal, frameTransform.FlipVertical, null, null,
            OfficeTextOverflowBehavior.Clip, textAdvanceWidth ?? width,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, OfficeTextBaseline.Normal, allowOverflow: false);

    /// <summary>Adds a transformed source text run inside a drawing-local clip path, retaining the resolved glyph advance and complete source text.</summary>
    public OfficeDrawing AddClippedPositionedText(
        string text, double x, double y, double width, double height,
        double clipX, double clipY, OfficeClipPath clipPath,
        OfficeImageFrameTransform frameTransform,
        OfficeFontInfo? font = null, OfficeColor? color = null,
        OfficeTextAlignment alignment = OfficeTextAlignment.Left,
        double? lineHeight = null, double? textAdvanceWidth = null) {
        if (clipPath == null) throw new ArgumentNullException(nameof(clipPath));
        ValidateFiniteNonNegative(clipX, nameof(clipX));
        ValidateFiniteNonNegative(clipY, nameof(clipY));
        if (clipX + clipPath.Width > Width || clipY + clipPath.Height > Height)
            throw new ArgumentOutOfRangeException(nameof(clipPath), "Text clip must fit inside the drawing bounds.");
        var clipped = new OfficeDrawing(Math.Max(0.01D, clipPath.Width), Math.Max(0.01D, clipPath.Height));
        clipped.AddTextCore(text, x - clipX, y - clipY, width, height, font, color, alignment, lineHeight,
            OfficeTextVerticalAlignment.Top, frameTransform.RotationDegrees, frameTransform.CenterX - clipX, frameTransform.CenterY - clipY,
            false, false, false, frameTransform.FlipHorizontal, frameTransform.FlipVertical, null, null,
            OfficeTextOverflowBehavior.Clip, textAdvanceWidth ?? width,
            OfficeTextDecorationStyle.None, OfficeTextDecorationStyle.None, OfficeTextBaseline.Normal, allowOverflow: true);
        return AddClippedDrawing(clipped, clipX, clipY, clipPath);
    }
}
