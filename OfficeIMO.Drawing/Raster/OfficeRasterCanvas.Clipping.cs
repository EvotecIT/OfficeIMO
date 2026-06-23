using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private OfficeRasterClipRectangle? _clipRectangle;

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
        OfficeRasterClipRectangle? previous = _clipRectangle;
        OfficeRasterClipRectangle next = OfficeRasterClipRectangle.FromBounds(x, y, width, height);
        _clipRectangle = previous.HasValue
            ? OfficeRasterClipRectangle.Intersect(previous.Value, next)
            : next;
        return new ClipScope(this, previous);
    }

    private bool IsPixelInsideClip(int x, int y) =>
        !_clipRectangle.HasValue || _clipRectangle.Value.Contains(x, y);

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

    private sealed class ClipScope : IDisposable {
        private readonly OfficeRasterCanvas _canvas;
        private readonly OfficeRasterClipRectangle? _previous;
        private bool _disposed;

        internal ClipScope(OfficeRasterCanvas canvas, OfficeRasterClipRectangle? previous) {
            _canvas = canvas;
            _previous = previous;
        }

        public void Dispose() {
            if (_disposed) {
                return;
            }

            _canvas._clipRectangle = _previous;
            _disposed = true;
        }
    }
}
