using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

internal static partial class PdfWriter {
    private sealed partial class LayoutContext {
        /// <summary>Proves that all nominal paint lies within every applicable simple rectangular clip.</summary>
        private bool DiagnosticPaintContained(DiagnosticPaintBounds bounds,
            List<OfficeTextInkClip> sourceClips, List<OfficeTextInkClip> targetClips,
            OfficeTextInkClip frameClip, OfficeTextInkClip pageClip,
            TableCellContentFrame? secondClip, OfficeTransform transform) {
            if (sourceClips.Count + targetClips.Count + (secondClip.HasValue ? 3 : 2) > 64) return false;
            if (!bounds.IsFinite) return false;
            if (!DiagnosticClipContains(frameClip, bounds) || !DiagnosticClipContains(pageClip, bounds)) return false;
            foreach (OfficeTextInkClip clip in sourceClips) if (!DiagnosticClipContains(clip, bounds)) return false;
            if (!DiagnosticClipsContain(targetClips, bounds)) return false;
            return !secondClip.HasValue || DiagnosticClipContains(DiagnosticRectangleClip(secondClip.Value, transform), bounds);
        }

        /// <summary>Encloses the actual contour points; non-finite or empty paint has no usable bounds.</summary>
        private DiagnosticPaintBounds MeasureDiagnosticPaintBounds(List<List<OfficePoint>> contours) {
            double left = double.PositiveInfinity, top = double.PositiveInfinity;
            double right = double.NegativeInfinity, bottom = double.NegativeInfinity;
            foreach (List<OfficePoint> contour in contours) {
                cancellationToken.ThrowIfCancellationRequested();
                for (int index = 0; index < contour.Count; index++) {
                    if ((index & 1023) == 0) cancellationToken.ThrowIfCancellationRequested();
                    OfficePoint point = contour[index];
                    if (double.IsNaN(point.X) || double.IsInfinity(point.X)
                        || double.IsNaN(point.Y) || double.IsInfinity(point.Y)) return default;
                    left = Math.Min(left, point.X); top = Math.Min(top, point.Y);
                    right = Math.Max(right, point.X); bottom = Math.Max(bottom, point.Y);
                }
            }
            return left > right || top > bottom ? default : new DiagnosticPaintBounds(left, top, right, bottom);
        }

        /// <summary>
        /// A clip that encloses every contour point removes no ink, so adding it cannot cause loss.
        /// This is the same proof that lets fully contained paint skip ink measurement.
        /// </summary>
        private bool DiagnosticClipsContain(List<OfficeTextInkClip> clips, DiagnosticPaintBounds bounds) {
            foreach (OfficeTextInkClip clip in clips) if (!DiagnosticClipContains(clip, bounds)) return false;
            return true;
        }

        private bool DiagnosticClipContains(OfficeTextInkClip clip, DiagnosticPaintBounds bounds) {
            cancellationToken.ThrowIfCancellationRequested();
            return bounds.IsFinite && clip.FilledContours != null
                && OfficeTextInkClip.TryGetAxisAlignedRectangle(clip.FilledContours,
                    out double clipLeft, out double clipTop, out double clipRight, out double clipBottom)
                && Finite(clipLeft) && Finite(clipTop) && Finite(clipRight) && Finite(clipBottom)
                && bounds.Left >= clipLeft && bounds.Top >= clipTop && bounds.Right <= clipRight && bounds.Bottom <= clipBottom;

            static bool Finite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);
        }

        private readonly struct DiagnosticPaintBounds {
            internal DiagnosticPaintBounds(double left, double top, double right, double bottom) {
                Left = left; Top = top; Right = right; Bottom = bottom; IsFinite = true;
            }
            internal double Left { get; }
            internal double Top { get; }
            internal double Right { get; }
            internal double Bottom { get; }
            internal bool IsFinite { get; }
        }
    }
}
