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
    private readonly List<OfficeDrawingElement> _elements = new List<OfficeDrawingElement>();
    private readonly ReadOnlyCollection<OfficeDrawingElement> _elementsView;

    /// <summary>Drawing width in the caller's layout unit.</summary>
    public double Width { get; }

    /// <summary>Drawing height in the caller's layout unit.</summary>
    public double Height { get; }

    /// <summary>Positioned shapes in paint order.</summary>
    public IReadOnlyList<OfficeDrawingShape> Shapes => _shapesView;

    /// <summary>Positioned drawing elements in paint order.</summary>
    public IReadOnlyList<OfficeDrawingElement> Elements => _elementsView;

    /// <summary>Creates a drawing canvas.</summary>
    public OfficeDrawing(double width, double height) {
        ValidatePositiveFinite(width, nameof(width));
        ValidatePositiveFinite(height, nameof(height));

        Width = width;
        Height = height;
        _shapesView = new ReadOnlyCollection<OfficeDrawingShape>(_shapes);
        _elementsView = new ReadOnlyCollection<OfficeDrawingElement>(_elements);
    }

    /// <summary>Adds a shape at a local top-left coordinate and returns this drawing.</summary>
    public OfficeDrawing AddShape(OfficeShape shape, double x, double y) {
        var item = new OfficeDrawingShape(shape, x, y);
        if (item.X + item.Shape.Width > Width || item.Y + item.Shape.Height > Height) {
            throw new ArgumentOutOfRangeException(nameof(shape), "Drawing shapes must fit inside the drawing bounds.");
        }

        _shapes.Add(item);
        _elements.Add(item);
        return this;
    }

    /// <summary>Adds text inside a local drawing rectangle and returns this drawing.</summary>
    public OfficeDrawing AddText(string text, double x, double y, double width, double height, OfficeFontInfo? font = null, OfficeColor? color = null, OfficeTextAlignment alignment = OfficeTextAlignment.Left, double? lineHeight = null, OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top, double rotationDegrees = 0D, double? rotationCenterX = null, double? rotationCenterY = null, bool wrapText = false) {
        var item = new OfficeDrawingText(text, x, y, width, height, font, color, alignment, lineHeight, verticalAlignment, rotationDegrees, rotationCenterX, rotationCenterY, wrapText);
        if (item.X + item.Width > Width || item.Y + item.Height > Height) {
            throw new ArgumentOutOfRangeException(nameof(text), "Drawing text must fit inside the drawing bounds.");
        }

        _elements.Add(item);
        return this;
    }

    /// <summary>Creates a detached copy of this drawing and all positioned elements.</summary>
    public OfficeDrawing Clone() {
        var clone = new OfficeDrawing(Width, Height);
        for (int i = 0; i < _elements.Count; i++) {
            OfficeDrawingElement element = _elements[i].CloneElement();
            clone._elements.Add(element);
            if (element is OfficeDrawingShape shape) {
                clone._shapes.Add(shape);
            }
        }

        return clone;
    }

    private static void ValidatePositiveFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0) {
            throw new ArgumentOutOfRangeException(paramName, "Drawing dimensions must be finite positive numbers.");
        }
    }
}
