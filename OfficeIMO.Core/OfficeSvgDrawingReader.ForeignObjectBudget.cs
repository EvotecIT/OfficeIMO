using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static bool TryMeasureForeignObjectSurfaces(OfficeDrawing drawing, int maximumElements,
        out int elementCount, out double pixels) {
        elementCount = 0;
        pixels = 0D;
        var pending = new Stack<OfficeDrawing>();
        pending.Push(drawing);
        while (pending.Count > 0) {
            foreach (OfficeDrawingElement element in pending.Pop().Elements) {
                if (++elementCount > maximumElements) return false;
                if (element is OfficeDrawingEffectGroup effect) {
                    if (effect.Opacity <= 0D) continue;
                    OfficeDrawing inner = effect.InnerDrawing;
                    double surface = Math.Ceiling(inner.Width) * Math.Ceiling(inner.Height);
                    pixels += surface * (effect.SoftMask == null ? 1D : 4D);
                    pending.Push(inner);
                    if (effect.SoftMask != null) pending.Push(effect.SoftMask.InnerDrawing);
                } else if (element is OfficeDrawingGroup group) {
                    if (group.ClipPath.Kind != OfficeClipPathKind.Empty) pending.Push(group.InnerDrawing);
                } else if (element is OfficeDrawingTilingPattern pattern) {
                    if (pattern.Opacity <= 0D) continue;
                    OfficeDrawing tile = pattern.InnerTile;
                    pixels += Math.Ceiling(tile.Width) * Math.Ceiling(tile.Height);
                    pending.Push(tile);
                } else if (element is OfficeDrawingImage image) {
                    if (!TryMeasureEncodedImage(image.EncodedBytes, image.Opacity, out double imagePixels)) return false;
                    pixels += imagePixels;
                } else if (element is OfficeDrawingImagePattern imagePattern) {
                    if (!TryMeasureEncodedImage(imagePattern.EncodedBytes, imagePattern.Opacity, out double patternPixels)) return false;
                    pixels += patternPixels;
                }
                if (double.IsNaN(pixels) || double.IsInfinity(pixels)) return false;
            }
        }
        return true;
    }

    private static bool TryMeasureEncodedImage(byte[] bytes, double opacity, out double pixels) {
        pixels = 0D;
        if (!OfficeImageReader.TryIdentifyByContent(bytes, null, out OfficeImageInfo info)
            || info.Format == OfficeImageFormat.Svg
            || !OfficeRasterGuards.TryEnsurePixelCount(info.Width, info.Height, out _)) return false;
        pixels = (double)info.Width * info.Height * (opacity < 1D ? 2D : 1D);
        return true;
    }
}
