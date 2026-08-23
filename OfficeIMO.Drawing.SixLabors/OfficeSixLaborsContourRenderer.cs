using System;
using System.Collections.Generic;
using System.Numerics;
using System.Threading;
using OfficeIMO.Drawing;
using SixLabors.Fonts;
using SixLabors.Fonts.Rendering;

namespace OfficeIMO.Drawing.SixLabors;

internal sealed class OfficeSixLaborsContourRenderer : IGlyphRenderer {
    private const double CurveTolerance = 0.05D;
    private const int MaximumCurveDepth = 12;
    private const int MaximumArcSegments = 512;
    private readonly List<List<OfficePoint>> _contours = new();
    private readonly int _maximumPointCount;
    private readonly CancellationToken _cancellationToken;
    private List<OfficePoint>? _current;
    private Vector2 _currentPoint;
    private int _pointCount;

    internal OfficeSixLaborsContourRenderer(
        int maximumPointCount,
        CancellationToken cancellationToken) {
        if (maximumPointCount <= 0) throw new ArgumentOutOfRangeException(nameof(maximumPointCount));
        _maximumPointCount = maximumPointCount;
        _cancellationToken = cancellationToken;
    }

    public void BeginText(in FontRectangle bounds) {
    }

    public void EndText() => EndFigure();

    public bool BeginGlyph(in FontRectangle bounds, in GlyphRendererParameters parameters) => true;

    public void EndGlyph() => EndFigure();

    public void BeginLayer(Paint? paint, FillRule fillRule) {
    }

    public void EndLayer() {
    }

    public void BeginGroup(CompositeMode mode) {
    }

    public void EndGroup() {
    }

    public void BeginFigure() => EndFigure();

    public void MoveTo(Vector2 point) {
        EndFigure();
        _current = new List<OfficePoint>();
        Add(point);
    }

    public void LineTo(Vector2 point) => Add(point);

    public void QuadraticBezierTo(Vector2 secondControlPoint, Vector2 point) {
        EnsureFigure();
        FlattenQuadratic(_currentPoint, secondControlPoint, point, 0);
    }

    public void CubicBezierTo(Vector2 secondControlPoint, Vector2 thirdControlPoint, Vector2 point) {
        EnsureFigure();
        FlattenCubic(_currentPoint, secondControlPoint, thirdControlPoint, point, 0);
    }

    public void ArcTo(
        float radiusX,
        float radiusY,
        float rotation,
        bool largeArc,
        bool sweep,
        Vector2 point) {
        EnsureFigure();
        if (radiusX <= 0F || radiusY <= 0F || point == _currentPoint) {
            Add(point);
            return;
        }
        AppendSvgArc(_currentPoint, point, radiusX, radiusY, rotation, largeArc, sweep);
    }

    public void EndFigure() {
        if (_current == null) return;
        if (_current.Count >= 3) {
            OfficePoint first = _current[0];
            OfficePoint last = _current[_current.Count - 1];
            if (first != last) {
                _cancellationToken.ThrowIfCancellationRequested();
                if (_pointCount >= _maximumPointCount) {
                    throw new InvalidOperationException("Font outline expansion exceeded the configured point budget.");
                }
                _current.Add(first);
                _pointCount++;
            }
            _contours.Add(_current);
        }
        _current = null;
    }

    public TextDecorations EnabledDecorations() => TextDecorations.None;

    public void SetDecoration(
        TextDecorations textDecorations,
        Vector2 start,
        Vector2 end,
        float thickness,
        ReadOnlyMemory<float> intersections) {
    }

    internal List<List<OfficePoint>> GetContours() {
        EndFigure();
        return _contours;
    }

    private void EnsureFigure() {
        if (_current == null) {
            _current = new List<OfficePoint>();
            Add(_currentPoint);
        }
    }

    private void Add(Vector2 point) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (_pointCount >= _maximumPointCount) {
            throw new InvalidOperationException("Font outline expansion exceeded the configured point budget.");
        }
        EnsureFigureCore();
        _current!.Add(new OfficePoint(point.X, point.Y));
        _pointCount++;
        _currentPoint = point;
    }

    private void EnsureFigureCore() {
        if (_current == null) _current = new List<OfficePoint>();
    }

    private void AppendSvgArc(
        Vector2 start,
        Vector2 end,
        float rawRadiusX,
        float rawRadiusY,
        float rotationDegrees,
        bool largeArc,
        bool sweep) {
        double rx = Math.Abs(rawRadiusX);
        double ry = Math.Abs(rawRadiusY);
        double phi = rotationDegrees * Math.PI / 180D;
        double cosPhi = Math.Cos(phi);
        double sinPhi = Math.Sin(phi);
        double dx = (start.X - end.X) / 2D;
        double dy = (start.Y - end.Y) / 2D;
        double xPrime = (cosPhi * dx) + (sinPhi * dy);
        double yPrime = (-sinPhi * dx) + (cosPhi * dy);
        double radiiScale = (xPrime * xPrime / (rx * rx)) + (yPrime * yPrime / (ry * ry));
        if (radiiScale > 1D) {
            double scale = Math.Sqrt(radiiScale);
            rx *= scale;
            ry *= scale;
        }

        double rx2 = rx * rx;
        double ry2 = ry * ry;
        double numerator = Math.Max(0D, (rx2 * ry2) - (rx2 * yPrime * yPrime) - (ry2 * xPrime * xPrime));
        double denominator = (rx2 * yPrime * yPrime) + (ry2 * xPrime * xPrime);
        double coefficient = denominator <= double.Epsilon
            ? 0D
            : (largeArc == sweep ? -1D : 1D) * Math.Sqrt(numerator / denominator);
        double centerPrimeX = coefficient * (rx * yPrime / ry);
        double centerPrimeY = coefficient * (-ry * xPrime / rx);
        double centerX = (cosPhi * centerPrimeX) - (sinPhi * centerPrimeY) + ((start.X + end.X) / 2D);
        double centerY = (sinPhi * centerPrimeX) + (cosPhi * centerPrimeY) + ((start.Y + end.Y) / 2D);

        double startAngle = VectorAngle(1D, 0D, (xPrime - centerPrimeX) / rx, (yPrime - centerPrimeY) / ry);
        double sweepAngle = VectorAngle(
            (xPrime - centerPrimeX) / rx,
            (yPrime - centerPrimeY) / ry,
            (-xPrime - centerPrimeX) / rx,
            (-yPrime - centerPrimeY) / ry);
        if (!sweep && sweepAngle > 0D) sweepAngle -= 2D * Math.PI;
        if (sweep && sweepAngle < 0D) sweepAngle += 2D * Math.PI;
        double maximumRadius = Math.Max(rx, ry);
        double maximumAngle = maximumRadius <= CurveTolerance
            ? Math.PI / 6D
            : 2D * Math.Acos(Math.Max(-1D, Math.Min(1D, 1D - (CurveTolerance / maximumRadius))));
        if (double.IsNaN(maximumAngle) || maximumAngle <= 0D) maximumAngle = Math.PI / 24D;
        int segments = Math.Max(1, Math.Min(
            MaximumArcSegments,
            (int)Math.Ceiling(Math.Abs(sweepAngle) / maximumAngle)));
        for (int index = 1; index <= segments; index++) {
            double angle = startAngle + (sweepAngle * index / segments);
            double cosine = Math.Cos(angle);
            double sine = Math.Sin(angle);
            Add(new Vector2(
                (float)(centerX + (cosPhi * rx * cosine) - (sinPhi * ry * sine)),
                (float)(centerY + (sinPhi * rx * cosine) + (cosPhi * ry * sine))));
        }
    }

    private void FlattenQuadratic(Vector2 start, Vector2 control, Vector2 end, int depth) {
        if (depth >= MaximumCurveDepth || DistanceToLine(control, start, end) <= CurveTolerance) {
            Add(end);
            return;
        }

        Vector2 startControl = (start + control) / 2F;
        Vector2 controlEnd = (control + end) / 2F;
        Vector2 midpoint = (startControl + controlEnd) / 2F;
        FlattenQuadratic(start, startControl, midpoint, depth + 1);
        FlattenQuadratic(midpoint, controlEnd, end, depth + 1);
    }

    private void FlattenCubic(
        Vector2 start,
        Vector2 control1,
        Vector2 control2,
        Vector2 end,
        int depth) {
        double flatness = Math.Max(
            DistanceToLine(control1, start, end),
            DistanceToLine(control2, start, end));
        if (depth >= MaximumCurveDepth || flatness <= CurveTolerance) {
            Add(end);
            return;
        }

        Vector2 startControl = (start + control1) / 2F;
        Vector2 controls = (control1 + control2) / 2F;
        Vector2 controlEnd = (control2 + end) / 2F;
        Vector2 leftControl = (startControl + controls) / 2F;
        Vector2 rightControl = (controls + controlEnd) / 2F;
        Vector2 midpoint = (leftControl + rightControl) / 2F;
        FlattenCubic(start, startControl, leftControl, midpoint, depth + 1);
        FlattenCubic(midpoint, rightControl, controlEnd, end, depth + 1);
    }

    private static double DistanceToLine(Vector2 point, Vector2 start, Vector2 end) {
        double deltaX = end.X - start.X;
        double deltaY = end.Y - start.Y;
        double length = Math.Sqrt((deltaX * deltaX) + (deltaY * deltaY));
        if (length <= double.Epsilon) return Vector2.Distance(point, start);
        return Math.Abs((deltaY * point.X) - (deltaX * point.Y) + (end.X * start.Y) - (end.Y * start.X)) / length;
    }

    private static double VectorAngle(double ux, double uy, double vx, double vy) {
        double denominator = Math.Sqrt(((ux * ux) + (uy * uy)) * ((vx * vx) + (vy * vy)));
        if (denominator <= double.Epsilon) return 0D;
        double cosine = Math.Max(-1D, Math.Min(1D, ((ux * vx) + (uy * vy)) / denominator));
        double angle = Math.Acos(cosine);
        return (ux * vy) - (uy * vx) < 0D ? -angle : angle;
    }
}
