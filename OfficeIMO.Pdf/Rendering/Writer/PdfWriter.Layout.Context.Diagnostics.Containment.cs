using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Proves that all nominal paint lies within every applicable simple rectangular clip.</summary>
        private bool DiagnosticPaintContained(List<List<OfficePoint>> contours,
            List<OfficeTextInkClip> sourceClips, List<OfficeTextInkClip> targetClips,
            OfficeTextInkClip frameClip, OfficeTextInkClip pageClip,
            TableCellContentFrame? secondClip, OfficeTransform transform) {
            if (sourceClips.Count + targetClips.Count + (secondClip.HasValue ? 3 : 2) > 64) return false;
            double left = double.PositiveInfinity, top = double.PositiveInfinity;
            double right = double.NegativeInfinity, bottom = double.NegativeInfinity;
            foreach (List<OfficePoint> contour in contours) {
                cancellationToken.ThrowIfCancellationRequested();
                for (int index = 0; index < contour.Count; index++) {
                    if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                    OfficePoint point = contour[index];
                    if (double.IsNaN(point.X) || double.IsInfinity(point.X)
                        || double.IsNaN(point.Y) || double.IsInfinity(point.Y)) return false;
                    left = Math.Min(left, point.X); top = Math.Min(top, point.Y);
                    right = Math.Max(right, point.X); bottom = Math.Max(bottom, point.Y);
                }
            }
            if (left > right || top > bottom) return false;
            if (!Contains(frameClip) || !Contains(pageClip)) return false;
            foreach (OfficeTextInkClip clip in sourceClips) if (!Contains(clip)) return false;
            foreach (OfficeTextInkClip clip in targetClips) if (!Contains(clip)) return false;
            return !secondClip.HasValue || Contains(DiagnosticRectangleClip(secondClip.Value, transform));

            bool Contains(OfficeTextInkClip clip) {
                cancellationToken.ThrowIfCancellationRequested();
                return clip.FilledContours != null
                    && OfficeTextInkClip.TryGetAxisAlignedRectangle(clip.FilledContours,
                        out double clipLeft, out double clipTop, out double clipRight, out double clipBottom)
                    && Finite(clipLeft) && Finite(clipTop) && Finite(clipRight) && Finite(clipBottom)
                    && left >= clipLeft && top >= clipTop && right <= clipRight && bottom <= clipBottom;
            }

            static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
        }
    }
}
