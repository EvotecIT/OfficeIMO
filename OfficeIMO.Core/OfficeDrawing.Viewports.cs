using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeDrawing {
    /// <summary>
    /// Retains overflow visible through a fitted viewport, including isolated effect canvases.
    /// Explicit clipping paths remain unchanged. Used by importers whose viewBox is not a clip.
    /// </summary>
    internal bool TryExpandViewportCanvas(double visibleLeft, double visibleTop, double visibleRight, double visibleBottom,
        double maximumDimension, double maximumPixels, out OfficeDrawing expanded, out double left, out double top,
        int depth = 0) {
        expanded = this;
        left = Math.Min(0D, visibleLeft);
        top = Math.Min(0D, visibleTop);
        double right = Math.Max(Width, visibleRight);
        double bottom = Math.Max(Height, visibleBottom);
        double width = right - left, height = bottom - top;
        if (depth > 128 || double.IsNaN(width) || double.IsNaN(height) || width <= 0D || height <= 0D ||
            width > maximumDimension || height > maximumDimension || width * height > maximumPixels) return false;
        OfficeDrawing content = Clone();
        for (int index = 0; index < content._elements.Count; index++) {
            OfficeDrawingElement element = content._elements[index];
            if (element is OfficeDrawingEffectGroup effect && effect.Transform.TryInvert(out OfficeTransform inverse)) {
                var visible = inverse.TransformRectangleBounds(left, top, width, height);
                if (!effect.InnerDrawing.TryExpandViewportCanvas(visible.Left, visible.Top, visible.Right, visible.Bottom,
                        maximumDimension, maximumPixels, out OfficeDrawing child, out double childLeft, out double childTop,
                        depth + 1)) return false;
                OfficeDrawingSoftMask? mask = effect.SoftMask;
                if (mask != null && mask.Transform.TryInvert(out OfficeTransform maskInverse)) {
                    var maskVisible = maskInverse.TransformRectangleBounds(visible.Left, visible.Top,
                        visible.Right - visible.Left, visible.Bottom - visible.Top);
                    if (!mask.InnerDrawing.TryExpandViewportCanvas(maskVisible.Left, maskVisible.Top, maskVisible.Right, maskVisible.Bottom,
                            maximumDimension, maximumPixels, out OfficeDrawing maskDrawing, out double maskLeft, out double maskTop,
                            depth + 1)) return false;
                    mask = new OfficeDrawingSoftMask(maskDrawing, mask.Mode,
                        OfficeTransform.Translate(maskLeft, maskTop).Then(mask.Transform)
                            .Then(OfficeTransform.Translate(-childLeft, -childTop)),
                        mask.BackdropColor, mask.LuminosityStandard);
                }
                content.ReplaceElement(index, element, new OfficeDrawingEffectGroup(child,
                    OfficeTransform.Translate(childLeft, childTop).Then(effect.Transform), effect.BlendMode, mask, effect.Opacity));
            } else if (element is OfficeDrawingGroup group) {
                OfficeTransform frame = group.FrameTransform?.CreateDestinationTransform() ?? OfficeTransform.Identity;
                if (!frame.TryInvert(out OfficeTransform inverseFrame)) continue;
                var visible = inverseFrame.TransformRectangleBounds(left, top, width, height);
                double clipLeft = Math.Max(visible.Left, group.X);
                double clipTop = Math.Max(visible.Top, group.Y);
                double clipRight = Math.Min(visible.Right, group.X + group.ClipPath.Width);
                double clipBottom = Math.Min(visible.Bottom, group.Y + group.ClipPath.Height);
                if (clipRight <= clipLeft || clipBottom <= clipTop) continue;
                double offsetX = group.X + group.ContentOffsetX, offsetY = group.Y + group.ContentOffsetY;
                if (!group.InnerDrawing.TryExpandViewportCanvas(clipLeft - offsetX, clipTop - offsetY,
                        clipRight - offsetX, clipBottom - offsetY, maximumDimension, maximumPixels,
                        out OfficeDrawing child, out double childLeft, out double childTop, depth + 1)) return false;
                content.ReplaceElement(index, element, new OfficeDrawingGroup(child, group.X, group.Y, group.ClipPath,
                    group.ContentOffsetX + childLeft, group.ContentOffsetY + childTop, group.FrameTransform,
                    group.ActualText, group.ActualTextAnchorX, group.ActualTextAnchorY));
            }
        }
        expanded = new OfficeDrawing(width, height) { Fonts = Fonts.Clone() };
        expanded.AddDrawingForClippedRendering(content, -left, -top, null);
        return true;
    }
}
