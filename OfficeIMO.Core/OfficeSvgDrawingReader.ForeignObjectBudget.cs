using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public static partial class OfficeSvgDrawingReader {
    private static bool TryMeasureForeignObjectSurfaces(OfficeDrawing drawing, int maximumElements,
        out int elementCount, out double pixels) {
        elementCount = 0;
        pixels = 0D;
        var pending = new Stack<(OfficeDrawing Drawing, bool FrameTransform)>();
        pending.Push((drawing, false));
        while (pending.Count > 0) {
            (OfficeDrawing current, bool inheritedFrameTransform) = pending.Pop();
            foreach (OfficeDrawingElement element in current.Elements) {
                if (++elementCount > maximumElements) return false;
                if (element is OfficeDrawingEffectGroup effect) {
                    if (effect.Opacity <= 0D) continue;
                    OfficeDrawing inner = effect.InnerDrawing;
                    double surface = Math.Ceiling(inner.Width) * Math.Ceiling(inner.Height);
                    // The effect layer and ApplySoftMask's mask-scene/result buffers
                    // use three source-sized surfaces; the nested mask uses its own size.
                    pixels += effect.SoftMask == null ? surface : surface * 3D +
                        Math.Ceiling(effect.SoftMask.InnerDrawing.Width) * Math.Ceiling(effect.SoftMask.InnerDrawing.Height);
                    // A parent frame transforms the completed effect surfaces,
                    // not text while its inner scene is being rendered.
                    pending.Push((inner, false));
                    if (effect.SoftMask != null) pending.Push((effect.SoftMask.InnerDrawing, false));
                } else if (element is OfficeDrawingGroup group) {
                    if (group.ClipPath.Kind != OfficeClipPathKind.Empty) pending.Push((group.InnerDrawing,
                        inheritedFrameTransform || group.FrameTransform.HasValue && group.FrameTransform.Value.HasTransform));
                } else if (element is OfficeDrawingTilingPattern pattern) {
                    if (pattern.Opacity <= 0D) continue;
                    OfficeDrawing tile = pattern.InnerTile;
                    pixels += Math.Ceiling(tile.Width) * Math.Ceiling(tile.Height);
                    // Tiling likewise transforms its finished tile surface.
                    pending.Push((tile, false));
                } else if (element is OfficeDrawingImage image) {
                    if (!TryMeasureEncodedImage(image.EncodedBytes, image.Opacity, out double imagePixels)) return false;
                    pixels += imagePixels;
                } else if (element is OfficeDrawingImagePattern imagePattern) {
                    if (!TryMeasureEncodedImage(imagePattern.EncodedBytes, imagePattern.Opacity, out double patternPixels)) return false;
                    pixels += patternPixels;
                } else if (element is OfficeDrawingText text && (text.HasFrameTransform || inheritedFrameTransform)) {
                    // A positioned or vertical text layer can be larger than its
                    // declared frame once the font is shaped. The callback has no
                    // renderer-independent bound for that intermediate surface.
                    return false;
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
