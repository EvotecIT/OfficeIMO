using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>Intent associated with an ink stroke.</summary>
public enum OfficeInkBias {
    /// <summary>The stroke is intended as handwriting.</summary>
    Handwriting = 0,
    /// <summary>The stroke is intended as a drawing.</summary>
    Drawing = 1,
    /// <summary>The stroke may be interpreted as handwriting or drawing.</summary>
    Both = 2
}

/// <summary>Shape of the pen tip used to paint an ink stroke.</summary>
public enum OfficeInkTipShape {
    /// <summary>An elliptical pen tip.</summary>
    Ellipse = 0,
    /// <summary>A rectangular pen tip.</summary>
    Rectangle = 1
}

/// <summary>A sampled point in a format-neutral ink stroke.</summary>
public readonly struct OfficeInkPoint : IEquatable<OfficeInkPoint> {
    /// <summary>Horizontal coordinate in the owning canvas unit.</summary>
    public double X { get; }

    /// <summary>Vertical coordinate in the owning canvas unit.</summary>
    public double Y { get; }

    /// <summary>Optional normalized pressure from 0 through 1.</summary>
    public double? Pressure { get; }

    /// <summary>Optional horizontal pen tilt in degrees.</summary>
    public double? TiltX { get; }

    /// <summary>Optional vertical pen tilt in degrees.</summary>
    public double? TiltY { get; }

    /// <summary>Optional sample timestamp relative to the stroke origin.</summary>
    public TimeSpan? Timestamp { get; }

    /// <summary>Creates an ink sample.</summary>
    public OfficeInkPoint(
        double x,
        double y,
        double? pressure = null,
        double? tiltX = null,
        double? tiltY = null,
        TimeSpan? timestamp = null) {
        ValidateFinite(x, nameof(x));
        ValidateFinite(y, nameof(y));
        ValidateOptionalFinite(pressure, nameof(pressure));
        ValidateOptionalFinite(tiltX, nameof(tiltX));
        ValidateOptionalFinite(tiltY, nameof(tiltY));
        if (pressure.HasValue && (pressure.Value < 0D || pressure.Value > 1D)) {
            throw new ArgumentOutOfRangeException(nameof(pressure), "Ink pressure must be from 0 through 1.");
        }
        if (timestamp.HasValue && timestamp.Value < TimeSpan.Zero) {
            throw new ArgumentOutOfRangeException(nameof(timestamp), "Ink sample time cannot be negative.");
        }

        X = x;
        Y = y;
        Pressure = pressure;
        TiltX = tiltX;
        TiltY = tiltY;
        Timestamp = timestamp;
    }

    /// <summary>Returns this sample transformed into another canvas coordinate space.</summary>
    public OfficeInkPoint Transform(OfficeTransform transform) {
        OfficePoint point = transform.TransformPoint(new OfficePoint(X, Y));
        return new OfficeInkPoint(point.X, point.Y, Pressure, TiltX, TiltY, Timestamp);
    }

    /// <inheritdoc />
    public bool Equals(OfficeInkPoint other) =>
        X.Equals(other.X) && Y.Equals(other.Y) &&
        Nullable.Equals(Pressure, other.Pressure) &&
        Nullable.Equals(TiltX, other.TiltX) &&
        Nullable.Equals(TiltY, other.TiltY) &&
        Nullable.Equals(Timestamp, other.Timestamp);

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficeInkPoint other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = X.GetHashCode();
            hash = (hash * 397) ^ Y.GetHashCode();
            hash = (hash * 397) ^ (Pressure?.GetHashCode() ?? 0);
            hash = (hash * 397) ^ (TiltX?.GetHashCode() ?? 0);
            hash = (hash * 397) ^ (TiltY?.GetHashCode() ?? 0);
            hash = (hash * 397) ^ (Timestamp?.GetHashCode() ?? 0);
            return hash;
        }
    }

    /// <summary>Returns true when two samples are equal.</summary>
    public static bool operator ==(OfficeInkPoint left, OfficeInkPoint right) => left.Equals(right);

    /// <summary>Returns true when two samples are different.</summary>
    public static bool operator !=(OfficeInkPoint left, OfficeInkPoint right) => !left.Equals(right);

    private static void ValidateFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, "Ink coordinates must be finite numbers.");
        }
    }

    private static void ValidateOptionalFinite(double? value, string paramName) {
        if (value.HasValue) ValidateFinite(value.Value, paramName);
    }
}

/// <summary>Axis-aligned bounds of format-neutral ink content.</summary>
public readonly struct OfficeInkBounds : IEquatable<OfficeInkBounds> {
    /// <summary>Horizontal origin.</summary>
    public double X { get; }

    /// <summary>Vertical origin.</summary>
    public double Y { get; }

    /// <summary>Bounds width.</summary>
    public double Width { get; }

    /// <summary>Bounds height.</summary>
    public double Height { get; }

    /// <summary>Whether the bounds contain no ink samples.</summary>
    public bool IsEmpty { get; }

    private OfficeInkBounds(double x, double y, double width, double height, bool isEmpty) {
        X = x;
        Y = y;
        Width = width;
        Height = height;
        IsEmpty = isEmpty;
    }

    /// <summary>Creates non-empty ink bounds.</summary>
    public OfficeInkBounds(double x, double y, double width, double height)
        : this(x, y, width, height, false) {
        if (double.IsNaN(x) || double.IsInfinity(x)) throw new ArgumentOutOfRangeException(nameof(x));
        if (double.IsNaN(y) || double.IsInfinity(y)) throw new ArgumentOutOfRangeException(nameof(y));
        if (double.IsNaN(width) || double.IsInfinity(width) || width < 0D) throw new ArgumentOutOfRangeException(nameof(width));
        if (double.IsNaN(height) || double.IsInfinity(height) || height < 0D) throw new ArgumentOutOfRangeException(nameof(height));
    }

    /// <summary>An empty bounds value.</summary>
    public static OfficeInkBounds Empty => new OfficeInkBounds(0D, 0D, 0D, 0D, true);

    /// <summary>Returns bounds covering this value and another value.</summary>
    public OfficeInkBounds Union(OfficeInkBounds other) {
        if (IsEmpty) return other;
        if (other.IsEmpty) return this;
        double left = Math.Min(X, other.X);
        double top = Math.Min(Y, other.Y);
        double right = Math.Max(X + Width, other.X + other.Width);
        double bottom = Math.Max(Y + Height, other.Y + other.Height);
        return new OfficeInkBounds(left, top, right - left, bottom - top);
    }

    /// <inheritdoc />
    public bool Equals(OfficeInkBounds other) =>
        X.Equals(other.X) && Y.Equals(other.Y) && Width.Equals(other.Width) &&
        Height.Equals(other.Height) && IsEmpty == other.IsEmpty;

    /// <inheritdoc />
    public override bool Equals(object? obj) => obj is OfficeInkBounds other && Equals(other);

    /// <inheritdoc />
    public override int GetHashCode() {
        unchecked {
            int hash = X.GetHashCode();
            hash = (hash * 397) ^ Y.GetHashCode();
            hash = (hash * 397) ^ Width.GetHashCode();
            hash = (hash * 397) ^ Height.GetHashCode();
            hash = (hash * 397) ^ IsEmpty.GetHashCode();
            return hash;
        }
    }
}

/// <summary>A dependency-free, format-neutral ink stroke.</summary>
public sealed class OfficeInkStroke {
    private readonly List<OfficeInkPoint> _points = new List<OfficeInkPoint>();
    private readonly ReadOnlyCollection<OfficeInkPoint> _pointsView;

    /// <summary>Creates an empty stroke.</summary>
    public OfficeInkStroke() {
        _pointsView = new ReadOnlyCollection<OfficeInkPoint>(_points);
    }

    /// <summary>Creates a stroke from sampled points.</summary>
    public OfficeInkStroke(IEnumerable<OfficeInkPoint> points) : this() {
        if (points == null) throw new ArgumentNullException(nameof(points));
        _points.AddRange(points);
    }

    /// <summary>Stroke samples in capture order.</summary>
    public IReadOnlyList<OfficeInkPoint> Points => _pointsView;

    /// <summary>Base stroke color.</summary>
    public OfficeColor Color { get; set; } = OfficeColor.Black;

    /// <summary>Nominal pen-tip width in canvas units.</summary>
    public double Width { get; set; } = 1.5D;

    /// <summary>Nominal pen-tip height in canvas units.</summary>
    public double Height { get; set; } = 1.5D;

    /// <summary>Additional stroke opacity from 0 through 1.</summary>
    public double Opacity { get; set; } = 1D;

    /// <summary>Pen tip shape.</summary>
    public OfficeInkTipShape TipShape { get; set; } = OfficeInkTipShape.Ellipse;

    /// <summary>Handwriting/drawing interpretation bias.</summary>
    public OfficeInkBias Bias { get; set; } = OfficeInkBias.Both;

    /// <summary>Whether a renderer should smooth the captured samples.</summary>
    public bool FitToCurve { get; set; }

    /// <summary>Whether pressure samples should be ignored.</summary>
    public bool IgnorePressure { get; set; }

    /// <summary>Whether the stroke represents translucent highlighting.</summary>
    public bool IsHighlighter { get; set; }

    /// <summary>Optional affine transform applied to captured samples.</summary>
    public OfficeTransform? Transform { get; set; }

    /// <summary>Optional handwriting language identifier.</summary>
    public uint? LanguageId { get; set; }

    /// <summary>Best recognized text associated with the stroke.</summary>
    public string? RecognizedText { get; set; }

    /// <summary>Alternative handwriting-recognition candidates, best first.</summary>
    public IList<string> RecognitionAlternatives { get; } = new List<string>();

    /// <summary>Adds a sampled point and returns this stroke.</summary>
    public OfficeInkStroke AddPoint(OfficeInkPoint point) {
        _points.Add(point);
        return this;
    }

    /// <summary>Adds a sampled point and returns this stroke.</summary>
    public OfficeInkStroke AddPoint(double x, double y, double? pressure = null) =>
        AddPoint(new OfficeInkPoint(x, y, pressure));

    /// <summary>Returns axis-aligned bounds including nominal pen-tip dimensions.</summary>
    public OfficeInkBounds GetBounds() {
        if (_points.Count == 0) return OfficeInkBounds.Empty;
        ValidateStyle();
        OfficeTransform transform = Transform ?? OfficeTransform.Identity;
        OfficePoint first = transform.TransformPoint(new OfficePoint(_points[0].X, _points[0].Y));
        double left = first.X;
        double top = first.Y;
        double right = first.X;
        double bottom = first.Y;
        for (int index = 1; index < _points.Count; index++) {
            OfficePoint point = transform.TransformPoint(new OfficePoint(_points[index].X, _points[index].Y));
            left = Math.Min(left, point.X);
            top = Math.Min(top, point.Y);
            right = Math.Max(right, point.X);
            bottom = Math.Max(bottom, point.Y);
        }
        double scale = Transform.HasValue ? TransformScale(Transform.Value) : 1D;
        double halfWidth = Width * scale / 2D;
        double halfHeight = Height * scale / 2D;
        return new OfficeInkBounds(left - halfWidth, top - halfHeight, right - left + halfWidth * 2D, bottom - top + halfHeight * 2D);
    }

    /// <summary>Creates a detached copy of this stroke.</summary>
    public OfficeInkStroke Clone() {
        var clone = new OfficeInkStroke(_points) {
            Color = Color,
            Width = Width,
            Height = Height,
            Opacity = Opacity,
            TipShape = TipShape,
            Bias = Bias,
            FitToCurve = FitToCurve,
            IgnorePressure = IgnorePressure,
            IsHighlighter = IsHighlighter,
            Transform = Transform,
            LanguageId = LanguageId,
            RecognizedText = RecognizedText
        };
        foreach (string alternative in RecognitionAlternatives) clone.RecognitionAlternatives.Add(alternative);
        return clone;
    }

    internal void ValidateStyle() {
        if (double.IsNaN(Width) || double.IsInfinity(Width) || Width <= 0D) throw new InvalidOperationException("Ink stroke width must be finite and positive.");
        if (double.IsNaN(Height) || double.IsInfinity(Height) || Height <= 0D) throw new InvalidOperationException("Ink stroke height must be finite and positive.");
        if (double.IsNaN(Opacity) || double.IsInfinity(Opacity) || Opacity < 0D || Opacity > 1D) throw new InvalidOperationException("Ink stroke opacity must be from 0 through 1.");
    }

    internal static double TransformScale(OfficeTransform transform) {
        double scaleX = Math.Sqrt(transform.M11 * transform.M11 + transform.M12 * transform.M12);
        double scaleY = Math.Sqrt(transform.M21 * transform.M21 + transform.M22 * transform.M22);
        return Math.Max(0.000001D, (scaleX + scaleY) / 2D);
    }
}

/// <summary>A reusable collection of ink strokes in one canvas coordinate system.</summary>
public sealed class OfficeInkDocument {
    private readonly List<OfficeInkStroke> _strokes = new List<OfficeInkStroke>();
    private readonly ReadOnlyCollection<OfficeInkStroke> _strokesView;

    /// <summary>Creates an empty ink document.</summary>
    public OfficeInkDocument() {
        _strokesView = new ReadOnlyCollection<OfficeInkStroke>(_strokes);
    }

    /// <summary>Ink strokes in paint order.</summary>
    public IReadOnlyList<OfficeInkStroke> Strokes => _strokesView;

    /// <summary>Adds a stroke and returns this document.</summary>
    public OfficeInkDocument Add(OfficeInkStroke stroke) {
        if (stroke == null) throw new ArgumentNullException(nameof(stroke));
        _strokes.Add(stroke);
        return this;
    }

    /// <summary>Returns bounds covering every stroke.</summary>
    public OfficeInkBounds GetBounds() {
        OfficeInkBounds bounds = OfficeInkBounds.Empty;
        for (int index = 0; index < _strokes.Count; index++) bounds = bounds.Union(_strokes[index].GetBounds());
        return bounds;
    }

    /// <summary>Creates a detached copy of this ink document.</summary>
    public OfficeInkDocument Clone() {
        var clone = new OfficeInkDocument();
        for (int index = 0; index < _strokes.Count; index++) clone.Add(_strokes[index].Clone());
        return clone;
    }
}
