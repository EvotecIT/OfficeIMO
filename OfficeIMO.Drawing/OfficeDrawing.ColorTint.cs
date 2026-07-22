using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeDrawing {
    /// <summary>
    /// Recolors paint-producing vector elements while preserving their original alpha.
    /// Used by uncolored pattern consumers after cloning a reusable tile scene.
    /// </summary>
    internal void ApplyColorTint(OfficeColor tint) {
        for (int i = 0; i < _elements.Count; i++) {
            OfficeDrawingElement current = _elements[i];
            OfficeDrawingElement? replacement = null;
            if (current is OfficeDrawingShape drawingShape) {
                OfficeShape shape = drawingShape.Shape;
                if (shape.FillColor.HasValue && shape.FillColor.Value.A > 0) shape.FillColor = WithTint(shape.FillColor.Value, tint);
                if (shape.StrokeColor.HasValue && shape.StrokeColor.Value.A > 0) shape.StrokeColor = WithTint(shape.StrokeColor.Value, tint);
            } else if (current is OfficeDrawingText text) {
                replacement = new OfficeDrawingText(
                    text.Text, text.X, text.Y, text.Width, text.Height, text.Font,
                    WithTint(text.Color ?? OfficeColor.Black, tint), text.Alignment, text.LineHeight,
                    text.VerticalAlignment, text.RotationDegrees, text.RotationCenterX, text.RotationCenterY,
                    text.WrapText, text.ShrinkToFit, text.StackedText, text.FlipHorizontal, text.FlipVertical,
                    text.Padding, text.ParagraphIndent, text.OverflowBehavior, text.TextAdvanceWidth);
            } else if (current is OfficeDrawingRichText richText) {
                var runs = new List<OfficeRichTextRun>(richText.Runs.Count);
                for (int runIndex = 0; runIndex < richText.Runs.Count; runIndex++) {
                    OfficeRichTextRun run = richText.Runs[runIndex];
                    runs.Add(new OfficeRichTextRun(
                        run.Text, run.FontSize, WithTint(run.Color, tint), run.Bold, run.Italic, run.Underline,
                        run.FontFamily, run.Strikethrough,
                        run.BackgroundColor.HasValue ? WithTint(run.BackgroundColor.Value, tint) : null));
                }
                replacement = new OfficeDrawingRichText(
                    runs, richText.X, richText.Y, richText.Width, richText.Height, richText.Alignment,
                    richText.LineHeight, richText.VerticalAlignment, richText.RotationDegrees,
                    richText.RotationCenterX, richText.RotationCenterY, richText.WrapText, richText.ShrinkToFit,
                    richText.FlipHorizontal, richText.FlipVertical, richText.Padding, richText.ParagraphIndent);
            } else if (current is OfficeDrawingGroup group) {
                OfficeDrawing child = group.InnerDrawing.Clone();
                child.ApplyColorTint(tint);
                replacement = new OfficeDrawingGroup(child, group.X, group.Y, group.ClipPath, group.ContentOffsetX, group.ContentOffsetY, group.FrameTransform);
            } else if (current is OfficeDrawingEffectGroup effectGroup) {
                OfficeDrawing child = effectGroup.InnerDrawing.Clone();
                child.ApplyColorTint(tint);
                replacement = new OfficeDrawingEffectGroup(child, effectGroup.Transform, effectGroup.BlendMode, effectGroup.SoftMask, effectGroup.Opacity);
            } else if (current is OfficeDrawingTilingPattern pattern) {
                OfficeDrawing tile = pattern.InnerTile.Clone();
                tile.ApplyColorTint(tint);
                replacement = new OfficeDrawingTilingPattern(
                    tile, pattern.Area, pattern.HorizontalStep, pattern.VerticalStep,
                    pattern.RepeatX, pattern.RepeatY, pattern.Transform,
                    pattern.OriginX, pattern.OriginY, pattern.MaximumTileCount, pattern.Opacity);
            }

            if (replacement != null) ReplaceElement(i, current, replacement);
        }
    }

    private void ReplaceElement(int index, OfficeDrawingElement current, OfficeDrawingElement replacement) {
        bool behindContent = _behindContentElements.Remove(current);
        _elements[index] = replacement;
        if (behindContent) _behindContentElements.Add(replacement);
    }

    private static OfficeColor WithTint(OfficeColor source, OfficeColor tint) =>
        OfficeColor.FromRgba(tint.R, tint.G, tint.B, source.A);
}
