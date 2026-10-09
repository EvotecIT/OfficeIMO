using System;

namespace OfficeIMO.Drawing;

public sealed partial class OfficeRasterCanvas {
    private bool IntersectsVisibleImageBounds((double Left, double Top, double Right, double Bottom) bounds, bool antialiasBoundary) {
        if (!antialiasBoundary) return IntersectsVisibleBounds(bounds);
        if (double.IsNaN(bounds.Left) || double.IsNaN(bounds.Top) ||
            double.IsNaN(bounds.Right) || double.IsNaN(bounds.Bottom)) return true;

        double left = .5D, top = .5D, right = Width - .5D, bottom = Height - .5D;
        if (_clipRegion != null) _clipRegion.IntersectBounds(ref left, ref top, ref right, ref bottom, _cancellationToken);
        // One visible pixel has equal minimum and maximum centre coordinates.
        // Its cell still has area; a surface touching only its outer edge does not.
        return right >= left && bottom >= top &&
            bounds.Right > left - .5D && bounds.Left < right + .5D &&
            bounds.Bottom > top - .5D && bounds.Top < bottom + .5D;
    }

    /// <summary>
    /// Measures the intersection of a destination pixel and the transformed source rectangle.
    /// Fully covered pixels retain their centre sample. Boundary pixels use their covered-area
    /// centroid and exact area, without adding a surface or supersampling interior pixels.
    /// </summary>
    private sealed class ImageBoundaryCoverage {
        private readonly OfficeTransform _inverse;
        private readonly double _width, _height, _radiusX, _radiusY;
        private OfficePoint[]? _polygon, _scratch;

        internal ImageBoundaryCoverage(OfficeTransform inverse, double width, double height) {
            _inverse = inverse;
            _width = width;
            _height = height;
            _radiusX = .5D * Math.Abs(inverse.M11) + .5D * Math.Abs(inverse.M21);
            _radiusY = .5D * Math.Abs(inverse.M12) + .5D * Math.Abs(inverse.M22);
        }

        internal double GetCoverage(OfficePoint centre, out OfficePoint sample) {
            sample = centre;
            if (!IsFinite(centre.X) || !IsFinite(centre.Y) ||
                centre.X + _radiusX <= 0D || centre.X - _radiusX >= _width ||
                centre.Y + _radiusY <= 0D || centre.Y - _radiusY >= _height) return 0D;
            if (centre.X - _radiusX >= 0D && centre.X + _radiusX <= _width &&
                centre.Y - _radiusY >= 0D && centre.Y + _radiusY <= _height) return 1D;

            // A rectangle clipped by another rectangle has at most eight vertices.
            // Allocate this fixed workspace only when a boundary pixel is reached.
            OfficePoint[] polygon = _polygon ??= new OfficePoint[8];
            OfficePoint[] scratch = _scratch ??= new OfficePoint[8];
            polygon[0] = new OfficePoint(-.5D, -.5D);
            polygon[1] = new OfficePoint(.5D, -.5D);
            polygon[2] = new OfficePoint(.5D, .5D);
            polygon[3] = new OfficePoint(-.5D, .5D);
            double scaleX = Math.Max(Math.Abs(_inverse.M11), Math.Abs(_inverse.M21));
            double scaleY = Math.Max(Math.Abs(_inverse.M12), Math.Abs(_inverse.M22));
            int count = ClipImagePixel(polygon, 4, scratch, _inverse.M11 / scaleX, _inverse.M21 / scaleX, centre.X / scaleX);
            count = ClipImagePixel(scratch, count, polygon, -_inverse.M11 / scaleX, -_inverse.M21 / scaleX, _width / scaleX - centre.X / scaleX);
            count = ClipImagePixel(polygon, count, scratch, _inverse.M12 / scaleY, _inverse.M22 / scaleY, centre.Y / scaleY);
            count = ClipImagePixel(scratch, count, polygon, -_inverse.M12 / scaleY, -_inverse.M22 / scaleY, _height / scaleY - centre.Y / scaleY);
            if (count < 3) return 0D;

            double twiceArea = 0D, centroidX = 0D, centroidY = 0D;
            OfficePoint previous = polygon[count - 1];
            for (int i = 0; i < count; i++) {
                OfficePoint current = polygon[i];
                double cross = previous.X * current.Y - current.X * previous.Y;
                twiceArea += cross;
                centroidX += (previous.X + current.X) * cross;
                centroidY += (previous.Y + current.Y) * cross;
                previous = current;
            }
            if (twiceArea <= 0D) return 0D;
            centroidX /= 3D * twiceArea;
            centroidY /= 3D * twiceArea;
            sample = new OfficePoint(
                Clamp(centre.X + _inverse.M11 * centroidX + _inverse.M21 * centroidY, 0D, _width),
                Clamp(centre.Y + _inverse.M12 * centroidX + _inverse.M22 * centroidY, 0D, _height));
            return Clamp(twiceArea * .5D, 0D, 1D);
        }

        private static int ClipImagePixel(OfficePoint[] input, int count, OfficePoint[] output, double axisX, double axisY, double offset) {
            if (count == 0) return 0;
            int outputCount = 0;
            OfficePoint previous = input[count - 1];
            double previousDistance = axisX * previous.X + axisY * previous.Y + offset;
            for (int i = 0; i < count; i++) {
                OfficePoint current = input[i];
                double distance = axisX * current.X + axisY * current.Y + offset;
                if ((distance >= 0D) != (previousDistance >= 0D) && distance != 0D && previousDistance != 0D) {
                    double fraction = previousDistance / (previousDistance - distance);
                    output[outputCount++] = new OfficePoint(
                        previous.X + fraction * (current.X - previous.X),
                        previous.Y + fraction * (current.Y - previous.Y));
                }
                if (distance >= 0D) output[outputCount++] = current;
                previous = current;
                previousDistance = distance;
            }
            return outputCount;
        }
    }
}
