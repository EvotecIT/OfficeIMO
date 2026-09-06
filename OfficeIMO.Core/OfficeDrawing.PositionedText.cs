using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeDrawing {
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
