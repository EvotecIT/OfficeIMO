using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free vector drawing canvas shared by OfficeIMO document packages.
/// Coordinates are expressed in the caller's layout unit and use a local top-left origin.
/// </summary>
public sealed class OfficeDrawing {
    private readonly List<OfficeDrawingShape> _shapes = new List<OfficeDrawingShape>();
    private readonly ReadOnlyCollection<OfficeDrawingShape> _shapesView;

    /// <summary>Drawing width in the caller's layout unit.</summary>
    public double Width { get; }

    /// <summary>Drawing height in the caller's layout unit.</summary>
    public double Height { get; }

    /// <summary>Positioned shapes in paint order.</summary>
    public IReadOnlyList<OfficeDrawingShape> Shapes => _shapesView;

    /// <summary>Creates a drawing canvas.</summary>
    public OfficeDrawing(double width, double height) {
        ValidatePositiveFinite(width, nameof(width));
        ValidatePositiveFinite(height, nameof(height));

        Width = width;
        Height = height;
        _shapesView = new ReadOnlyCollection<OfficeDrawingShape>(_shapes);
    }

    /// <summary>Adds a shape at a local top-left coordinate and returns this drawing.</summary>
    public OfficeDrawing AddShape(OfficeShape shape, double x, double y) {
        var item = new OfficeDrawingShape(shape, x, y);
        if (item.X + item.Shape.Width > Width || item.Y + item.Shape.Height > Height) {
            throw new ArgumentOutOfRangeException(nameof(shape), "Drawing shapes must fit inside the drawing bounds.");
        }

        _shapes.Add(item);
        return this;
    }

    /// <summary>Creates a detached copy of this drawing and all positioned shapes.</summary>
    public OfficeDrawing Clone() {
        var clone = new OfficeDrawing(Width, Height);
        for (int i = 0; i < _shapes.Count; i++) {
            clone._shapes.Add(_shapes[i].Clone());
        }

        return clone;
    }

    private static void ValidatePositiveFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0) {
            throw new ArgumentOutOfRangeException(paramName, "Drawing dimensions must be finite positive numbers.");
        }
    }
}
