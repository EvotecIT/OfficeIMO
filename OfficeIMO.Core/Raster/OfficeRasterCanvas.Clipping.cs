using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private OfficeRasterClipRegion? _clipRegion;

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
        OfficeRasterClipRectangle next = OfficeRasterClipRectangle.FromBounds(
            x * CoordinateScaleX, y * CoordinateScaleY, width * CoordinateScaleX, height * CoordinateScaleY);
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

        internal static OfficeRasterClipRectangle Intersect(OfficeRasterClipRectangle first, OfficeRasterClipRectangle second) =>
            new OfficeRasterClipRectangle(
                Math.Max(first.Left, second.Left),
                Math.Max(first.Top, second.Top),
                Math.Min(first.Right, second.Right),
                Math.Min(first.Bottom, second.Bottom));

        internal bool Contains(int x, int y) =>
            x >= Left && x < Right && y >= Top && y < Bottom;
    }

    private sealed class OfficeRasterClipRegion {
        private readonly OfficeRasterClipRectangle? _rectangle;
        private readonly IReadOnlyList<IReadOnlyList<OfficePoint>>? _contours;
        private readonly OfficeFillRule _fillRule;
        private readonly OfficeRasterClipRegion? _previous;

        private OfficeRasterClipRegion(OfficeRasterClipRectangle? rectangle, IReadOnlyList<IReadOnlyList<OfficePoint>>? contours, OfficeFillRule fillRule, OfficeRasterClipRegion? previous) {
            _rectangle = rectangle;
            _contours = contours;
            _fillRule = fillRule;
            _previous = previous;
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
