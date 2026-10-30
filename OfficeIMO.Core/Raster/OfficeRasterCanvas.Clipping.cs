using System;
using System.Collections.Generic;
using System.IO;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private OfficeRasterClipRegion? _clipRegion;
    private ClipWorkBudget _clipWorkBudget = new ClipWorkBudget();

    private sealed class ClipWorkBudget {
        internal long Maximum = 100_000_000L;
        internal long Consumed;
        internal void Charge(long edges) {
            if (edges > Maximum - Consumed) throw new InvalidDataException("Raster clip edge work budget exceeded.");
            Consumed += edges;
        }
    }

    internal long MaximumClipEdgeChecks { get => _clipWorkBudget.Maximum; set => _clipWorkBudget.Maximum = value; }

    /// <summary>Whether the current clip can be applied per source-over child without repeating fractional coverage.</summary>
    internal bool HasBinaryClipCoverage =>
        !AntialiasedClipping || _clipRegion == null || _clipRegion.HasBinaryCoverage();

    /// <summary>
    /// Restricts subsequent raster drawing to the supplied rectangular bounds until the returned scope is disposed.
    /// Nested clips are intersected with the current clip.
    /// </summary>
    /// <param name="x">Left edge of the clipping rectangle.</param>
    /// <param name="y">Top edge of the clipping rectangle.</param>
    /// <param name="width">Width of the clipping rectangle.</param>
    /// <param name="height">Height of the clipping rectangle.</param>
    /// <returns>A disposable scope that restores the previous clip.</returns>
    public IDisposable PushClipRectangle(double x, double y, double width, double height) {
        OfficeRasterClipRegion? previous = _clipRegion;
        OfficeRasterClipRectangle next = AntialiasedClipping
            ? OfficeRasterClipRectangle.FromFractionalBounds(x * CoordinateScaleX, y * CoordinateScaleY, width * CoordinateScaleX, height * CoordinateScaleY)
            : OfficeRasterClipRectangle.FromBounds(x * CoordinateScaleX, y * CoordinateScaleY, width * CoordinateScaleX, height * CoordinateScaleY);
        _clipRegion = OfficeRasterClipRegion.Rectangle(next, previous);
        return new ClipScope(this, previous);
    }

    /// <summary>
    /// Clips renderer-owned rectangular polygons using their original pixel-centre coverage.
    /// Unlike the public rectangular clip, this excludes boundary pixels whose centres lie outside the polygon.
    /// </summary>
    internal IDisposable PushClipRectangleAtPixelCentres(double left, double top, double right, double bottom) {
        left *= CoordinateScaleX; right *= CoordinateScaleX;
        top *= CoordinateScaleY; bottom *= CoordinateScaleY;
        OfficeRasterClipRegion? previous = _clipRegion;
        if (AntialiasedClipping) {
            _clipRegion = OfficeRasterClipRegion.Rectangle(
                OfficeRasterClipRectangle.FromFractionalBounds(left, top, right - left, bottom - top), previous);
            return new ClipScope(this, previous);
        }
        var next = new OfficeRasterClipRectangle(
            ResolvePixelCentreBoundary(left, Width),
            ResolvePixelCentreBoundary(top, Height),
            ResolvePixelCentreBoundary(right, Width),
            ResolvePixelCentreBoundary(bottom, Height));
        _clipRegion = OfficeRasterClipRegion.Rectangle(next, previous);
        return new ClipScope(this, previous);
    }

    private static int ResolvePixelCentreBoundary(double edge, int limit) {
        if (edge <= 0.5D) return 0;
        if (edge >= limit + 0.5D) return limit;
        return (int)Math.Ceiling(edge - 0.5D);
    }

    /// <summary>
    /// Restricts subsequent raster drawing to the supplied polygon until the returned scope is disposed.
    /// Nested clips are intersected with the current clip.
    /// </summary>
    /// <param name="points">Polygon points in canvas coordinates.</param>
    /// <returns>A disposable scope that restores the previous clip.</returns>
    public IDisposable PushClipPolygon(IReadOnlyList<OfficePoint> points) {
        OfficeRasterClipRegion? previous = _clipRegion;
        _clipRegion = OfficeRasterClipRegion.Polygon(ScaleCoordinates(points), previous);
        return new ClipScope(this, previous);
    }

    /// <summary>
    /// Restricts subsequent raster drawing to the supplied even-odd contours until the returned scope is disposed.
    /// Nested clips are intersected with the current clip.
    /// </summary>
    /// <param name="contours">Closed contours in canvas coordinates.</param>
    /// <returns>A disposable scope that restores the previous clip.</returns>
    public IDisposable PushClipPolygonsEvenOdd(IReadOnlyList<IReadOnlyList<OfficePoint>> contours) {
        OfficeRasterClipRegion? previous = _clipRegion;
        _clipRegion = OfficeRasterClipRegion.Polygons(ScaleCoordinates(contours), OfficeFillRule.EvenOdd, previous);
        return new ClipScope(this, previous);
    }

    /// <summary>
    /// Restricts subsequent raster drawing to the supplied non-zero winding contours until the returned scope is disposed.
    /// Nested clips are intersected with the current clip.
    /// </summary>
    /// <param name="contours">Closed contours in canvas coordinates.</param>
    /// <returns>A disposable scope that restores the previous clip.</returns>
    public IDisposable PushClipPolygonsNonZero(IReadOnlyList<IReadOnlyList<OfficePoint>> contours) {
        OfficeRasterClipRegion? previous = _clipRegion;
        _clipRegion = OfficeRasterClipRegion.Polygons(ScaleCoordinates(contours), OfficeFillRule.NonZero, previous);
        return new ClipScope(this, previous);
    }

    private bool IsPixelInsideClip(int x, int y) =>
        _clipRegion == null || _clipRegion.Contains(x, y);

    /// <summary>
    /// Anti-aliases clips: rectangles by their exact pixel-area overlap and polygons through an area-coverage
    /// mask built with the shared scanline filler.
    /// </summary>
    internal bool AntialiasedClipping { get; set; }

    /// <summary>Returns the clip coverage of a pixel in [0, 1].</summary>
    private double GetClipCoverage(int x, int y) {
        if (_clipRegion == null) return 1D;
        if (AntialiasedClipping) return _clipRegion.Coverage(this, x, y);
        if (_clipRegion.EdgeCount > 0) _clipWorkBudget.Charge(_clipRegion.EdgeCount);
        return _clipRegion.Contains(x, y) ? 1D : 0D;
    }

    /// <summary>
    /// Conservatively tests a complete affine surface against the physical canvas and current clip.
    /// Antialiased image boundaries retain pixels with partial area coverage; periodic tiles use pixel centres.
    /// </summary>
    internal bool IntersectsVisibleSurface(OfficeTransform transform, double width, double height, bool includePartialCoverage = true) =>
        IntersectsVisibleImageBounds(ScaleCoordinates(transform).TransformRectangleBounds(0D, 0D, width, height), includePartialCoverage);

    /// <summary>Tests current clip bounds against the canvas pixel centres; overlapping bounds remain inconclusive.</summary>
    internal bool HasVisibleClipBounds => IntersectsVisibleBounds((0D, 0D, Width, Height));

    private bool IntersectsVisibleBounds((double Left, double Top, double Right, double Bottom) bounds) {
        // Overflowed corner arithmetic is inconclusive, never proof of invisibility.
        if (double.IsNaN(bounds.Left) || double.IsNaN(bounds.Top) ||
            double.IsNaN(bounds.Right) || double.IsNaN(bounds.Bottom)) return true;
        // A reflected surface can include its maximum edge: the inverse maps that
        // edge to source coordinate zero. Preserve boundary pixel centres.
        double left = Math.Max(.5D, bounds.Left), top = Math.Max(.5D, bounds.Top);
        double right = Math.Min(Width - .5D, bounds.Right), bottom = Math.Min(Height - .5D, bounds.Bottom);
        if (_clipRegion != null) _clipRegion.IntersectBounds(ref left, ref top, ref right, ref bottom, _cancellationToken);
        return right >= left && bottom >= top;
    }

    private readonly struct OfficeRasterClipRectangle {
        internal OfficeRasterClipRectangle(int left, int top, int right, int bottom) {
            Left = left;
            Top = top;
            Right = right;
            Bottom = bottom;
        }

        internal int Left { get; }
        internal int Top { get; }
        internal int Right { get; }
        internal int Bottom { get; }

        internal static OfficeRasterClipRectangle FromBounds(double x, double y, double width, double height) {
            if (width <= 0D || height <= 0D || double.IsNaN(x) || double.IsNaN(y) || double.IsNaN(width) || double.IsNaN(height)) {
                return new OfficeRasterClipRectangle(0, 0, 0, 0);
            }

            int left = Clamp((int)Math.Floor(x), 0, int.MaxValue);
            int top = Clamp((int)Math.Floor(y), 0, int.MaxValue);
            int right = Clamp((int)Math.Ceiling(x + width), 0, int.MaxValue);
            int bottom = Clamp((int)Math.Ceiling(y + height), 0, int.MaxValue);
            return new OfficeRasterClipRectangle(left, top, Math.Max(left, right), Math.Max(top, bottom));
        }

        // Anti-aliased clips keep fractional edges; the integer range covers every partially visible pixel.
        internal static OfficeRasterClipRectangle FromFractionalBounds(double x, double y, double width, double height) {
            OfficeRasterClipRectangle outer = FromBounds(x, y, width, height);
            if (outer.Right <= outer.Left || outer.Bottom <= outer.Top) return outer;
            return new OfficeRasterClipRectangle(outer.Left, outer.Top, outer.Right, outer.Bottom, x, y, x + width, y + height);
        }

        private OfficeRasterClipRectangle(int left, int top, int right, int bottom, double edgeLeft, double edgeTop, double edgeRight, double edgeBottom)
            : this(left, top, right, bottom) {
            HasFractionalEdges = true;
            EdgeLeft = edgeLeft; EdgeTop = edgeTop; EdgeRight = edgeRight; EdgeBottom = edgeBottom;
        }

        internal bool HasFractionalEdges { get; }
        internal double EdgeLeft { get; }
        internal double EdgeTop { get; }
        internal double EdgeRight { get; }
        internal double EdgeBottom { get; }

        internal double Coverage(int x, int y) {
            if (!Contains(x, y)) return 0D;
            if (!HasFractionalEdges) return 1D;
            double horizontal = Math.Min(x + 1D, EdgeRight) - Math.Max(x, EdgeLeft);
            double vertical = Math.Min(y + 1D, EdgeBottom) - Math.Max(y, EdgeTop);
            return horizontal <= 0D || vertical <= 0D ? 0D : Math.Min(1D, horizontal) * Math.Min(1D, vertical);
        }

        internal static OfficeRasterClipRectangle Intersect(OfficeRasterClipRectangle first, OfficeRasterClipRectangle second) {
            int left = Math.Max(first.Left, second.Left), top = Math.Max(first.Top, second.Top);
            int right = Math.Min(first.Right, second.Right), bottom = Math.Min(first.Bottom, second.Bottom);
            if (!first.HasFractionalEdges && !second.HasFractionalEdges) return new OfficeRasterClipRectangle(left, top, right, bottom);
            double edgeLeft = Math.Max(first.HasFractionalEdges ? first.EdgeLeft : first.Left, second.HasFractionalEdges ? second.EdgeLeft : second.Left);
            double edgeTop = Math.Max(first.HasFractionalEdges ? first.EdgeTop : first.Top, second.HasFractionalEdges ? second.EdgeTop : second.Top);
            double edgeRight = Math.Min(first.HasFractionalEdges ? first.EdgeRight : first.Right, second.HasFractionalEdges ? second.EdgeRight : second.Right);
            double edgeBottom = Math.Min(first.HasFractionalEdges ? first.EdgeBottom : first.Bottom, second.HasFractionalEdges ? second.EdgeBottom : second.Bottom);
            return new OfficeRasterClipRectangle(left, top, Math.Max(left, right), Math.Max(top, bottom), edgeLeft, edgeTop, edgeRight, edgeBottom);
        }

        internal bool Contains(int x, int y) =>
            x >= Left && x < Right && y >= Top && y < Bottom;
    }

    private sealed partial class OfficeRasterClipRegion {
        private readonly OfficeRasterClipRectangle? _rectangle;
        private readonly IReadOnlyList<IReadOnlyList<OfficePoint>>? _contours;
        private readonly OfficeFillRule _fillRule;
        private readonly OfficeRasterClipRegion? _previous;
        internal long EdgeCount { get; }

        private OfficeRasterClipRegion(OfficeRasterClipRectangle? rectangle, IReadOnlyList<IReadOnlyList<OfficePoint>>? contours, OfficeFillRule fillRule, OfficeRasterClipRegion? previous) {
            _rectangle = rectangle;
            _contours = contours;
            _fillRule = fillRule;
            _previous = previous;
            long edges = previous?.EdgeCount ?? 0;
            if (contours != null) foreach (IReadOnlyList<OfficePoint> contour in contours)
                if (contour.Count >= 3) edges = checked(edges + contour.Count);
            EdgeCount = edges;
        }

        internal static OfficeRasterClipRegion Rectangle(OfficeRasterClipRectangle rectangle, OfficeRasterClipRegion? previous) {
            if (previous != null && previous._rectangle.HasValue) {
                rectangle = OfficeRasterClipRectangle.Intersect(rectangle, previous._rectangle.Value);
                previous = previous._previous;
            }
            return new OfficeRasterClipRegion(rectangle, null, OfficeFillRule.EvenOdd, previous);
        }

        internal static OfficeRasterClipRegion Polygon(IReadOnlyList<OfficePoint> points, OfficeRasterClipRegion? previous) =>
            new OfficeRasterClipRegion(null, new[] { points }, OfficeFillRule.EvenOdd, previous);

        internal static OfficeRasterClipRegion Polygons(IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeFillRule fillRule, OfficeRasterClipRegion? previous) =>
            new OfficeRasterClipRegion(null, contours, fillRule, previous);

        internal bool HasBinaryCoverage() {
            for (OfficeRasterClipRegion? current = this; current != null; current = current._previous) {
                OfficeRasterClipRectangle? rectangle = current._rectangle ?? current._maskRectangle;
                if (rectangle.HasValue) {
                    OfficeRasterClipRectangle bounds = rectangle.Value;
                    if (bounds.HasFractionalEdges && !HasIntegralEdges(bounds.EdgeLeft, bounds.EdgeTop, bounds.EdgeRight, bounds.EdgeBottom)) return false;
                } else if (current._contours == null ||
                    !OfficeTextInkClip.TryGetAxisAlignedRectangle(current._contours, out double left, out double top, out double right, out double bottom) ||
                    !HasIntegralEdges(left, top, right, bottom)) {
                    // Do not allocate a polygon mask merely to qualify an optimization.
                    return false;
                }
            }
            return true;
        }

        private static bool HasIntegralEdges(double left, double top, double right, double bottom) =>
            IsFinite(left) && IsFinite(top) && IsFinite(right) && IsFinite(bottom) &&
            left == Math.Truncate(left) && top == Math.Truncate(top) &&
            right == Math.Truncate(right) && bottom == Math.Truncate(bottom);

        internal void IntersectBounds(ref double left, ref double top, ref double right, ref double bottom,
            System.Threading.CancellationToken cancellationToken) {
            _previous?.IntersectBounds(ref left, ref top, ref right, ref bottom, cancellationToken);
            if (_rectangle.HasValue) {
                OfficeRasterClipRectangle rectangle = _rectangle.Value;
                // Rectangles store half-open integer pixel ranges, so an empty
                // range stays empty even when surface bounds meet at an edge.
                left = Math.Max(left, rectangle.Left + .5D); top = Math.Max(top, rectangle.Top + .5D);
                right = Math.Min(right, rectangle.Right - .5D); bottom = Math.Min(bottom, rectangle.Bottom - .5D);
            } else {
                double clipLeft = double.PositiveInfinity, clipTop = double.PositiveInfinity;
                double clipRight = double.NegativeInfinity, clipBottom = double.NegativeInfinity;
                if (_contours != null) foreach (IReadOnlyList<OfficePoint> contour in _contours) {
                    if (contour.Count < 3) continue;
                    foreach (OfficePoint point in contour) {
                        cancellationToken.ThrowIfCancellationRequested();
                        if (double.IsNaN(point.X) || double.IsNaN(point.Y)) return;
                        clipLeft = Math.Min(clipLeft, point.X); clipTop = Math.Min(clipTop, point.Y);
                        clipRight = Math.Max(clipRight, point.X); clipBottom = Math.Max(clipBottom, point.Y);
                    }
                }
                if (clipRight <= clipLeft || clipBottom <= clipTop) {
                    left = top = double.PositiveInfinity;
                    right = bottom = double.NegativeInfinity;
                    return;
                }
                left = Math.Max(left, clipLeft); top = Math.Max(top, clipTop);
                right = Math.Min(right, clipRight); bottom = Math.Min(bottom, clipBottom);
            }
        }

        internal bool Contains(int x, int y) {
            if (_previous != null && !_previous.Contains(x, y)) {
                return false;
            }

            if (_rectangle.HasValue) {
                return _rectangle.Value.Contains(x, y);
            }

            if (_contours == null || _contours.Count == 0) {
                return false;
            }

            double sampleX = x + 0.5D;
            double sampleY = y + 0.5D;
            int winding = 0;
            for (int i = 0; i < _contours.Count; i++) {
                IReadOnlyList<OfficePoint> contour = _contours[i];
                if (contour.Count < 3) {
                    continue;
                }

                if (_fillRule == OfficeFillRule.NonZero) {
                    winding += GetWindingNumber(contour, sampleX, sampleY);
                } else if (ContainsPoint(contour, sampleX, sampleY)) {
                    winding++;
                }
            }

            return _fillRule == OfficeFillRule.NonZero ? winding != 0 : (winding & 1) == 1;
        }
    }

    private sealed class ClipScope : IDisposable {
        private readonly OfficeRasterCanvas _canvas;
        private readonly OfficeRasterClipRegion? _previous;
        private bool _disposed;

        internal ClipScope(OfficeRasterCanvas canvas, OfficeRasterClipRegion? previous) {
            _canvas = canvas;
            _previous = previous;
        }

        public void Dispose() {
            if (_disposed) {
                return;
            }

            _canvas._clipRegion = _previous;
            _disposed = true;
        }
    }
}
