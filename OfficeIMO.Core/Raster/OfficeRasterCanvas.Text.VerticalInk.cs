using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private void InspectVerticalTextInk(OfficeDrawingText text, OfficeTransform placement,
        IReadOnlyList<OfficeTextInkClip> outerClips, Action charge,
        Action<(double Left, double Top, double Right, double Bottom, bool HasInk, bool IsMeasured, bool IsClipped), string?> report) {
        var (x, y, width, height) = OfficeDrawingRasterRenderer.ResolveTextContentRectangle(text, 1D);
        if (width <= 0D || height <= 0D) return;
        if (outerClips.Count >= 64) throw new NotSupportedException("Drawing text ink inspection exceeds its 64 nested groups/clips limit.");
        // The transformed painter shapes into a local layer before placing it.
        // Observe the same contours without allocating that raster intermediate.
        if (text.HasFrameTransform) {
            placement = OfficeDrawingRasterRenderer.CreateVerticalTextPlacement(text, 1D, x, y).Then(placement);
            x = 0D; y = 0D;
        }
        if (!placement.TryInvert(out OfficeTransform inverse)) {
            report((0D, 0D, 0D, 0D, false, false, false), "Vertical text placement cannot establish invertible ink geometry.");
            return;
        }
        var clips = new List<OfficeTextInkClip>(outerClips) {
            new OfficeTextInkClip(x, y, width, height, true, true, inverse)
        };
        bool preserve = PreservePaintedGlyphOrder;
        PreservePaintedGlyphOrder = text.PreservesPaintedGlyphs;
        using var metricScope = PushFontMetricScale(FontMetricScale * text.FontMetricScale);
        using var faceScope = PushTextFace(text.Font.Face);
        try {
            InspectLaidOutTextInk(() => {
                if (!TryDrawVerticalText(text.RasterText, x, y, width, height, text.Color ?? OfficeColor.Black,
                    text.Font.Size, text.Font.Style, text.Font.FamilyName, text.FeatureSettings, text.FontPalette,
                    text.UnderlineStyle, text.StrikethroughStyle, text.DecorationColor))
                    report((0D, 0D, 0D, 0D, false, false, false), "Vertical text has no usable shaped outlines; fallback layout is not inspected.");
            }, placement, clips, charge, report);
        } finally { PreservePaintedGlyphOrder = preserve; }
    }
}
