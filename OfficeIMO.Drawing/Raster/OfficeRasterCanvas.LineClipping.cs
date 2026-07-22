using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private bool TryClipLineToCanvas(
        ref OfficePoint start,
        ref OfficePoint end,
        double thickness,
        double originalLength,
        out double leadingDistance,
        out double trailingDistance) {
        double x1 = start.X;
        double y1 = start.Y;
        double x2 = end.X;
        double y2 = end.Y;
        bool visible = TryClipLineToCanvas(ref x1, ref y1, ref x2, ref y2, thickness, originalLength, out leadingDistance, out trailingDistance);
        if (visible) {
            start = new OfficePoint(x1, y1);
            end = new OfficePoint(x2, y2);
        }
        return visible;
    }

    private bool TryClipLineToCanvas(
        ref double x1,
        ref double y1,
        ref double x2,
        ref double y2,
        double thickness,
        double originalLength,
        out double leadingDistance,
        out double trailingDistance) {
        leadingDistance = 0D;
        trailingDistance = 0D;
        if (!IsFinite(x1) || !IsFinite(y1) || !IsFinite(x2) || !IsFinite(y2) || !IsFinite(originalLength) || originalLength <= 0D) return false;

        double originalX = x1;
        double originalY = y1;
        double dx = x2 - x1;
        double dy = y2 - y1;
        double padding = Math.Max(1D, thickness / 2D + 1D);
        double minimumX = -padding;
        double minimumY = -padding;
        double maximumX = Width - 1D + padding;
        double maximumY = Height - 1D + padding;
        double startRatio = 0D;
        double endRatio = 1D;
        if (!ClipLineBoundary(-dx, originalX - minimumX, ref startRatio, ref endRatio)
            || !ClipLineBoundary(dx, maximumX - originalX, ref startRatio, ref endRatio)
            || !ClipLineBoundary(-dy, originalY - minimumY, ref startRatio, ref endRatio)
            || !ClipLineBoundary(dy, maximumY - originalY, ref startRatio, ref endRatio)) return false;

        x1 = originalX + dx * startRatio;
        y1 = originalY + dy * startRatio;
        x2 = originalX + dx * endRatio;
        y2 = originalY + dy * endRatio;
        leadingDistance = originalLength * startRatio;
        trailingDistance = originalLength * (1D - endRatio);
        return true;
    }

    private static bool ClipLineBoundary(double direction, double distance, ref double startRatio, ref double endRatio) {
        if (Math.Abs(direction) <= double.Epsilon) return distance >= 0D;
        double ratio = distance / direction;
        if (direction < 0D) {
            if (ratio > endRatio) return false;
            if (ratio > startRatio) startRatio = ratio;
        } else {
            if (ratio < startRatio) return false;
            if (ratio < endRatio) endRatio = ratio;
        }
        return true;
    }

    private static double AdvancePatternPosition(double patternPosition, double distance, double cycle) {
        if (!IsFinite(cycle) || cycle <= 0D) return 0D;
        double normalizedPosition = IsFinite(patternPosition) ? patternPosition % cycle : 0D;
        double normalizedDistance = IsFinite(distance) ? distance % cycle : 0D;
        double distanceUntilWrap = cycle - normalizedPosition;
        double advanced = normalizedDistance >= distanceUntilWrap
            ? normalizedDistance - distanceUntilWrap
            : normalizedPosition + normalizedDistance;
        return advanced < 0D ? advanced + cycle : advanced;
    }
}
