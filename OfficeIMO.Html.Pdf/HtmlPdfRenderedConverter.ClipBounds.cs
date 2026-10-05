using OfficeIMO.Html;
using System;

namespace OfficeIMO.Html.Pdf;

internal static partial class HtmlPdfRenderedConverter {
    private readonly struct ClipBounds {
        internal ClipBounds(double left, double top, double right, double bottom, bool allowsInteractiveWidgets = true, bool constrainToSurface = true) {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
            AllowsInteractiveWidgets = allowsInteractiveWidgets;
            ConstrainToSurface = constrainToSurface;
        }

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

        internal ClipBounds Translate(double offsetX, double offsetY) => new(
            Left + offsetX, Top + offsetY, Right + offsetX, Bottom + offsetY,
            AllowsInteractiveWidgets, ConstrainToSurface);

        internal (double X, double Y, double Width, double Height) ConstrainLogicalRectangle(
            double x, double y, double width, double height) {
            // Keep contained carriers byte-for-byte unchanged. Subtracting x from
            // x + width can otherwise look like clipping through rounding alone.
            if (x >= Left - 0.0001D && y >= Top - 0.0001D
                && x + width <= Right + 0.0001D && y + height <= Bottom + 0.0001D) {
                return (x, y, width, height);
            }
            double left = Math.Max(x, Left), top = Math.Max(y, Top);
            double right = Math.Min(x + width, Right), bottom = Math.Min(y + height, Bottom);
            // Preserve paint and ownership. Only an intersecting carrier shrinks;
            // a wholly outside frame is not evidence of visible glyph ink.
            return right > left && bottom > top ? (left, top, right - left, bottom - top) : (x, y, width, height);
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
                active.Value.ConstrainToSurface && next.ConstrainToSurface);
    }

}
