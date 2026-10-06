using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    // Installed only on the inspection canvas. Null contours mean the painter
    // fell back to stroke text without a measurable font outline.
    private Action<IReadOnlyList<List<OfficePoint>>?>? _textInkObserver;

    // Decorations retain the inspection contract's conservative stroke envelope.
    private bool InspectTextDecoration(double x, double width, double y, double fontHeight,
        OfficeTextDecorationStyle style, double rotationRadians, double rotationCenterX, double rotationCenterY,
        bool flipHorizontal, bool flipVertical) {
        if (_textInkObserver == null) return false;
        double thickness = Math.Max(1D, fontHeight / 16D);
        double extent = thickness / 2D + (style == OfficeTextDecorationStyle.Wavy ? Math.Max(1D, thickness)
            : style == OfficeTextDecorationStyle.Double ? Math.Max(2D, thickness * 1.8D) / 2D : 0D);
        var contour = new List<OfficePoint> {
            new OfficePoint(x - thickness / 2D, y - extent), new OfficePoint(x + width + thickness / 2D, y - extent),
            new OfficePoint(x + width + thickness / 2D, y + extent), new OfficePoint(x - thickness / 2D, y + extent)
        };
        for (int i = 0; i < contour.Count; i++) contour[i] = TransformFramePoint(contour[i], rotationRadians,
            rotationCenterX, rotationCenterY, flipHorizontal, flipVertical);
        _textInkObserver(new[] { contour });
        return true;
    }

    private void InspectLaidOutTextInk(Action paint, OfficeTransform placement,
        IReadOnlyList<OfficeTextInkClip> clips, Action charge,
        Action<(double Left, double Top, double Right, double Bottom, bool HasInk, bool IsMeasured, bool IsClipped), string?> report) {
        var previous = _textInkObserver;
        long clipWork = 4_000_000;
        _textInkObserver = contours => {
            charge();
            if (contours == null) {
                report((0D, 0D, 0D, 0D, false, false, false), "Laid-out drawing text uses an unresolved font.");
                return;
            }
            bool clipped = false;
            var prepared = new List<List<OfficePoint>>(contours.Count);
            foreach (var contour in contours) {
                _cancellationToken.ThrowIfCancellationRequested();
                if ((clipWork -= contour.Count) < 0) throw new NotSupportedException("Laid-out text contour preparation exceeds its work limit.");
                var transformed = new List<OfficePoint>(contour.Count);
                foreach (var point in contour) transformed.Add(placement.TransformPoint(point));
                foreach (var clip in clips) if (clip.FilledContours == null)
                    transformed = clip.Apply(transformed, ref clipped, ref clipWork, _cancellationToken);
                prepared.Add(transformed);
            }
            var bounds = MeasureFilledContourBounds(prepared, OfficeFillRule.NonZero, clips);
            bounds.IsClipped |= clipped;
            report(bounds, bounds.IsMeasured ? null : "Laid-out drawing text exceeds supported contour geometry.");
        };
        try { paint(); } finally { _textInkObserver = previous; }
    }
}
