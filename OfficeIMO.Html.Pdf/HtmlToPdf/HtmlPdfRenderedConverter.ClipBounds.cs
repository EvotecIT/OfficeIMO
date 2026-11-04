using OfficeIMO.Html;
using OfficeIMO.Drawing;
using System;
using System.Collections.Generic;
using System.Linq;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private readonly struct ClipBounds {
        internal ClipBounds(double left, double top, double right, double bottom, bool allowsInteractiveWidgets = true, bool constrainToSurface = true)
            : this(left, top, right, bottom, allowsInteractiveWidgets, constrainToSurface, null) { }

        private ClipBounds(double left, double top, double right, double bottom, bool allowsInteractiveWidgets,
            bool constrainToSurface, CarrierWindow[]? carrierWindows) {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
            AllowsInteractiveWidgets = allowsInteractiveWidgets;
            ConstrainToSurface = constrainToSurface;
            _carrierWindows = carrierWindows;
        }

        private readonly CarrierWindow[]? _carrierWindows;
        private double Left { get; }
        private double Top { get; }
        private double Right { get; }
        private double Bottom { get; }
        internal bool AllowsInteractiveWidgets { get; }
        internal bool ConstrainToSurface { get; }
        internal static ClipBounds TransformedCoordinateSpace { get; } = new(
            double.NegativeInfinity,
            double.NegativeInfinity,
            double.PositiveInfinity,
            double.PositiveInfinity,
            allowsInteractiveWidgets: false,
            constrainToSurface: false);

        internal ClipBounds InEffectCoordinateSpace(OfficeTransform transform) {
            if (!transform.TryInvert(out OfficeTransform inverse)
                || double.IsInfinity(Left) || double.IsInfinity(Top)
                || double.IsInfinity(Right) || double.IsInfinity(Bottom)) {
                return TransformedCoordinateSpace;
            }
            var bounds = inverse.TransformRectangleBounds(Left, Top, Right - Left, Bottom - Top);
            // This window constrains logical carriers, not transformed glyph ink.
            // The graphics state retains the authored clip and physical page boundary.
            CarrierWindow[] windows = _carrierWindows ?? new[] { new CarrierWindow(Left, Top, Right, Bottom, OfficeTransform.Identity) };
            return new ClipBounds(bounds.Left, bounds.Top, bounds.Right, bounds.Bottom,
                false, false, windows.Select(window => window.InCoordinateSpace(transform)).ToArray());
        }

        internal ClipBounds Translate(double offsetX, double offsetY) => new(
            Left + offsetX, Top + offsetY, Right + offsetX, Bottom + offsetY,
            AllowsInteractiveWidgets, ConstrainToSurface,
            _carrierWindows?.Select(window => window.InCoordinateSpace(OfficeTransform.Translate(-offsetX, -offsetY))).ToArray());

        internal (double X, double Y, double Width, double Height) ConstrainLogicalRectangle(
            double x, double y, double width, double height) {
            // Keep contained carriers byte-for-byte unchanged. Subtracting x from
            // x + width can otherwise look like clipping through rounding alone.
            if (x >= Left - 0.0001D && y >= Top - 0.0001D
                && x + width <= Right + 0.0001D && y + height <= Bottom + 0.0001D
                && (_carrierWindows == null || _carrierWindows.All(window => window.Contains(x, y, width, height)))) {
                return (x, y, width, height);
            }
            double left = Math.Max(x, Left), top = Math.Max(y, Top);
            double right = Math.Min(x + width, Right), bottom = Math.Min(y + height, Bottom);
            var intersection = (Left: left, Top: top, Right: right, Bottom: bottom);
            if (_carrierWindows != null && right > left && bottom > top) {
                foreach (CarrierWindow window in _carrierWindows) {
                    window.ConstrainIndependentVerticalRange(ref top, ref bottom);
                }
                foreach (CarrierWindow window in _carrierWindows) {
                    window.ConstrainHorizontalRange(top, bottom, ref left, ref right);
                }
            }
            if (_carrierWindows != null && (right <= left || bottom <= top)
                && TryConstrainCornerIntersection(intersection, out var corner)) {
                return corner;
            }
            // Preserve paint and ownership. Only an intersecting carrier shrinks;
            // a wholly outside frame is not evidence of visible glyph ink.
            return right > left && bottom > top ? (left, top, right - left, bottom - top) : (x, y, width, height);
        }

        private bool TryConstrainCornerIntersection(
            (double Left, double Top, double Right, double Bottom) bounds,
            out (double X, double Y, double Width, double Height) rectangle) {
            rectangle = default;
            if (bounds.Right <= bounds.Left || bounds.Bottom <= bounds.Top) return false;
            var polygon = new List<(double X, double Y)> {
                (bounds.Left, bounds.Top), (bounds.Right, bounds.Top),
                (bounds.Right, bounds.Bottom), (bounds.Left, bounds.Bottom)
            };
            foreach (CarrierWindow window in _carrierWindows!) {
                polygon = window.Intersect(polygon);
                if (polygon.Count < 3) return false;
            }
            // A corner intersection may contain no full-height rectangle. Use an
            // interior point of its convex polygon and shrink only this fallback
            // carrier until all four corners satisfy every affine window.
            double centerX = polygon.Average(point => point.X), centerY = polygon.Average(point => point.Y);
            double halfWidth = Math.Min(centerX - bounds.Left, bounds.Right - centerX);
            double halfHeight = Math.Min(centerY - bounds.Top, bounds.Bottom - centerY);
            double scale = 1D;
            foreach (CarrierWindow window in _carrierWindows!) {
                scale = Math.Min(scale, window.GetContainedScale(centerX, centerY, halfWidth, halfHeight));
            }
            halfWidth *= scale;
            halfHeight *= scale;
            if (!(halfWidth > 0D && halfHeight > 0D)) return false;
            rectangle = (centerX - halfWidth, centerY - halfHeight, halfWidth * 2D, halfHeight * 2D);
            return true;
        }

        internal bool Contains(HtmlRenderVisual visual) {
            double right = visual.X + visual.Width;
            double bottom = visual.Y + visual.Height;
            return visual.X >= Left - 0.0001D && visual.Y >= Top - 0.0001D
                && right <= Right + 0.0001D && bottom <= Bottom + 0.0001D;
        }

        internal static ClipBounds Intersect(ClipBounds? active, ClipBounds next) => !active.HasValue
            ? next
            : new ClipBounds(
                Math.Max(active.Value.Left, next.Left),
                Math.Max(active.Value.Top, next.Top),
                Math.Min(active.Value.Right, next.Right),
                Math.Min(active.Value.Bottom, next.Bottom),
                active.Value.AllowsInteractiveWidgets && next.AllowsInteractiveWidgets,
                active.Value.ConstrainToSurface && next.ConstrainToSurface,
                MergeCarrierWindows(active.Value, next));

        private static CarrierWindow[]? MergeCarrierWindows(ClipBounds active, ClipBounds next) {
            if (active._carrierWindows == null && next._carrierWindows == null) return null;
            CarrierWindow[] first = active._carrierWindows ?? new[] { new CarrierWindow(active.Left, active.Top, active.Right, active.Bottom, OfficeTransform.Identity) };
            CarrierWindow[] second = next._carrierWindows ?? new[] { new CarrierWindow(next.Left, next.Top, next.Right, next.Bottom, OfficeTransform.Identity) };
            return first.Concat(second).ToArray();
        }

        private readonly struct CarrierWindow {
            private readonly double _left, _top, _right, _bottom;
            private readonly OfficeTransform _toWindow;

            internal CarrierWindow(double left, double top, double right, double bottom, OfficeTransform toWindow) {
                _left = left; _top = top; _right = right; _bottom = bottom; _toWindow = toWindow;
            }

            internal CarrierWindow InCoordinateSpace(OfficeTransform transform) =>
                new(_left, _top, _right, _bottom, transform.Then(_toWindow));

            internal bool Contains(double x, double y, double width, double height) {
                var bounds = _toWindow.TransformRectangleBounds(x, y, width, height);
                return bounds.Left >= _left - 0.0001D && bounds.Top >= _top - 0.0001D
                    && bounds.Right <= _right + 0.0001D && bounds.Bottom <= _bottom + 0.0001D;
            }

            internal void ConstrainIndependentVerticalRange(ref double top, ref double bottom) {
                if (Math.Abs(_toWindow.M11) < 0.000000000001D)
                    Constrain(_toWindow.M21, _toWindow.OffsetX, _left, _right, ref top, ref bottom);
                if (Math.Abs(_toWindow.M12) < 0.000000000001D)
                    Constrain(_toWindow.M22, _toWindow.OffsetY, _top, _bottom, ref top, ref bottom);
            }

            internal void ConstrainHorizontalRange(double top, double bottom, ref double left, ref double right) {
                // At both vertical endpoints, intersect the permitted x interval.
                // Linear affine boundaries then contain all four carrier corners.
                Constrain(_toWindow.M11, _toWindow.M21 * top + _toWindow.OffsetX, _left, _right, ref left, ref right);
                Constrain(_toWindow.M11, _toWindow.M21 * bottom + _toWindow.OffsetX, _left, _right, ref left, ref right);
                Constrain(_toWindow.M12, _toWindow.M22 * top + _toWindow.OffsetY, _top, _bottom, ref left, ref right);
                Constrain(_toWindow.M12, _toWindow.M22 * bottom + _toWindow.OffsetY, _top, _bottom, ref left, ref right);
            }

            internal List<(double X, double Y)> Intersect(List<(double X, double Y)> points) {
                points = ClipEdge(points, _toWindow.M11, _toWindow.M21, _right - _toWindow.OffsetX);
                points = ClipEdge(points, -_toWindow.M11, -_toWindow.M21, _toWindow.OffsetX - _left);
                points = ClipEdge(points, _toWindow.M12, _toWindow.M22, _bottom - _toWindow.OffsetY);
                return ClipEdge(points, -_toWindow.M12, -_toWindow.M22, _toWindow.OffsetY - _top);
            }

            internal double GetContainedScale(double x, double y, double halfWidth, double halfHeight) {
                double centerX = _toWindow.M11 * x + _toWindow.M21 * y + _toWindow.OffsetX;
                double centerY = _toWindow.M12 * x + _toWindow.M22 * y + _toWindow.OffsetY;
                double extentX = Math.Abs(_toWindow.M11) * halfWidth + Math.Abs(_toWindow.M21) * halfHeight;
                double extentY = Math.Abs(_toWindow.M12) * halfWidth + Math.Abs(_toWindow.M22) * halfHeight;
                double scaleX = extentX > 0D ? Math.Min(centerX - _left, _right - centerX) / extentX : 1D;
                double scaleY = extentY > 0D ? Math.Min(centerY - _top, _bottom - centerY) / extentY : 1D;
                return Math.Max(0D, Math.Min(scaleX, scaleY));
            }

            private static List<(double X, double Y)> ClipEdge(
                List<(double X, double Y)> input, double a, double b, double limit) {
                var output = new List<(double X, double Y)>();
                if (input.Count == 0) return output;
                var previous = input[input.Count - 1];
                double previousDistance = limit - a * previous.X - b * previous.Y;
                foreach (var point in input) {
                    double distance = limit - a * point.X - b * point.Y;
                    if ((distance >= 0D) != (previousDistance >= 0D)) {
                        double fraction = previousDistance / (previousDistance - distance);
                        output.Add((previous.X + (point.X - previous.X) * fraction,
                            previous.Y + (point.Y - previous.Y) * fraction));
                    }
                    if (distance >= 0D) output.Add(point);
                    previous = point;
                    previousDistance = distance;
                }
                return output;
            }

            private static void Constrain(double coefficient, double offset, double minimum, double maximum,
                ref double left, ref double right) {
                if (Math.Abs(coefficient) < 0.000000000001D) return;
                double first = (minimum - offset) / coefficient, second = (maximum - offset) / coefficient;
                left = Math.Max(left, Math.Min(first, second));
                right = Math.Min(right, Math.Max(first, second));
            }
        }
    }

}
