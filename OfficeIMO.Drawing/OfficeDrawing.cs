using System;
using System.Collections.Generic;
using System.Collections.ObjectModel;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free vector drawing canvas shared by OfficeIMO document packages.
/// Coordinates are expressed in the caller's layout unit and use a local top-left origin.
/// </summary>
public sealed partial class OfficeDrawing {
    private readonly List<OfficeDrawingShape> _shapes = new List<OfficeDrawingShape>();
    private readonly ReadOnlyCollection<OfficeDrawingShape> _shapesView;
    private readonly List<OfficeDrawingImage> _images = new List<OfficeDrawingImage>();
    private readonly ReadOnlyCollection<OfficeDrawingImage> _imagesView;
    private readonly List<OfficeDrawingElement> _elements = new List<OfficeDrawingElement>();
    private readonly ReadOnlyCollection<OfficeDrawingElement> _elementsView;

    /// <summary>Drawing width in the caller's layout unit.</summary>
    public double Width { get; }

    /// <summary>Drawing height in the caller's layout unit.</summary>
    public double Height { get; }

    /// <summary>Positioned shapes in paint order.</summary>
    public IReadOnlyList<OfficeDrawingShape> Shapes => _shapesView;

    /// <summary>Positioned images in paint order.</summary>
    public IReadOnlyList<OfficeDrawingImage> Images => _imagesView;

    /// <summary>Positioned drawing elements in paint order.</summary>
    public IReadOnlyList<OfficeDrawingElement> Elements => _elementsView;

    /// <summary>Creates a drawing canvas.</summary>
    public OfficeDrawing(double width, double height) {
        ValidatePositiveFinite(width, nameof(width));
        ValidatePositiveFinite(height, nameof(height));

        Width = width;
        Height = height;
        _shapesView = new ReadOnlyCollection<OfficeDrawingShape>(_shapes);
        _imagesView = new ReadOnlyCollection<OfficeDrawingImage>(_images);
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
    public OfficeDrawing AddText(string text, double x, double y, double width, double height, OfficeFontInfo? font = null, OfficeColor? color = null, OfficeTextAlignment alignment = OfficeTextAlignment.Left, double? lineHeight = null, OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top, double rotationDegrees = 0D, double? rotationCenterX = null, double? rotationCenterY = null, bool wrapText = false, bool shrinkToFit = false, bool stackedText = false, bool flipHorizontal = false, bool flipVertical = false, OfficeTextPadding? padding = null, OfficeTextParagraphIndent? paragraphIndent = null) {
        var item = new OfficeDrawingText(text, x, y, width, height, font, color, alignment, lineHeight, verticalAlignment, rotationDegrees, rotationCenterX, rotationCenterY, wrapText, shrinkToFit, stackedText, flipHorizontal, flipVertical, padding, paragraphIndent);
        if (item.X + item.Width > Width || item.Y + item.Height > Height) {
            throw new ArgumentOutOfRangeException(nameof(text), "Drawing text must fit inside the drawing bounds.");
        }

        _elements.Add(item);
        return this;
    }

    /// <summary>Adds rich text inside a local drawing rectangle and returns this drawing.</summary>
    public OfficeDrawing AddRichText(IReadOnlyList<OfficeRichTextRun> runs, double x, double y, double width, double height, OfficeTextAlignment alignment = OfficeTextAlignment.Left, double? lineHeight = null, OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top, double rotationDegrees = 0D, double? rotationCenterX = null, double? rotationCenterY = null, bool wrapText = true, bool shrinkToFit = false, bool flipHorizontal = false, bool flipVertical = false, OfficeTextPadding? padding = null, OfficeTextParagraphIndent? paragraphIndent = null) {
        var item = new OfficeDrawingRichText(runs, x, y, width, height, alignment, lineHeight, verticalAlignment, rotationDegrees, rotationCenterX, rotationCenterY, wrapText, shrinkToFit, flipHorizontal, flipVertical, padding, paragraphIndent);
        if (item.X + item.Width > Width || item.Y + item.Height > Height) {
            throw new ArgumentOutOfRangeException(nameof(runs), "Drawing rich text must fit inside the drawing bounds.");
        }

        _elements.Add(item);
        return this;
    }

    /// <summary>Adds an image using a shared placement/crop/transform projection and returns this drawing.</summary>
    public OfficeDrawing AddImage(byte[] bytes, string? contentType, OfficeImageProjection projection, string? alternativeText = null) {
        var item = new OfficeDrawingImage(bytes, contentType, projection, alternativeText);
        (double left, double top, double right, double bottom) = item.Projection.GetDestinationBounds();
        if (left < 0D || top < 0D || right > Width || bottom > Height) {
            throw new ArgumentOutOfRangeException(nameof(projection), "Drawing images must fit inside the drawing bounds.");
        }

        _images.Add(item);
        _elements.Add(item);
        return this;
    }

    /// <summary>Adds all elements from another drawing at a local destination offset and returns this drawing.</summary>
    public OfficeDrawing AddDrawing(OfficeDrawing drawing, double x, double y) {
        return AddDrawingCore(drawing, x, y, null);
    }

    /// <summary>Adds all elements from another drawing at a local destination offset with a shared frame transform.</summary>
    public OfficeDrawing AddDrawing(OfficeDrawing drawing, double x, double y, OfficeImageFrameTransform frameTransform) {
        return AddDrawingCore(drawing, x, y, frameTransform);
    }

    /// <summary>Adds another drawing as a clipped nested group at a local destination offset.</summary>
    public OfficeDrawing AddClippedDrawing(OfficeDrawing drawing, double x, double y, OfficeClipPath clipPath) {
        return AddClippedDrawingCore(drawing, x, y, clipPath, null);
    }

    /// <summary>Adds another drawing as a clipped nested group at a local destination offset with a shared frame transform.</summary>
    public OfficeDrawing AddClippedDrawing(OfficeDrawing drawing, double x, double y, OfficeClipPath clipPath, OfficeImageFrameTransform frameTransform) {
        return AddClippedDrawingCore(drawing, x, y, clipPath, frameTransform);
    }

    private OfficeDrawing AddClippedDrawingCore(OfficeDrawing drawing, double x, double y, OfficeClipPath clipPath, OfficeImageFrameTransform? frameTransform) {
        if (drawing == null) {
            throw new ArgumentNullException(nameof(drawing));
        }

        if (clipPath == null) {
            throw new ArgumentNullException(nameof(clipPath));
        }

        ValidateFiniteNonNegative(x, nameof(x));
        ValidateFiniteNonNegative(y, nameof(y));
        if (x + clipPath.Width > Width || y + clipPath.Height > Height) {
            throw new ArgumentOutOfRangeException(nameof(clipPath), "Nested drawing clip must fit inside the drawing bounds.");
        }

        _elements.Add(new OfficeDrawingGroup(drawing, x, y, clipPath, frameTransform));
        return this;
    }

    private OfficeDrawing AddDrawingCore(OfficeDrawing drawing, double x, double y, OfficeImageFrameTransform? frameTransform) {
        if (drawing == null) {
            throw new ArgumentNullException(nameof(drawing));
        }

        ValidateFiniteNonNegative(x, nameof(x));
        ValidateFiniteNonNegative(y, nameof(y));
        if (x + drawing.Width > Width || y + drawing.Height > Height) {
            throw new ArgumentOutOfRangeException(nameof(drawing), "Nested drawing content must fit inside the drawing bounds.");
        }

        for (int i = 0; i < drawing.Elements.Count; i++) {
            OfficeDrawingElement element = drawing.Elements[i];
            if (element is OfficeDrawingShape shape) {
                AddNestedShape(shape, x, y, frameTransform);
            } else if (element is OfficeDrawingText text) {
                AddNestedText(text, x, y, frameTransform);
            } else if (element is OfficeDrawingRichText richText) {
                AddNestedRichText(richText, x, y, frameTransform);
            } else if (element is OfficeDrawingImage image) {
                AddNestedImage(image, x, y, frameTransform);
            } else if (element is OfficeDrawingGroup group) {
                AddNestedGroup(group, x, y, frameTransform);
            }
        }

        return this;
    }

    private void AddNestedShape(OfficeDrawingShape drawingShape, double offsetX, double offsetY, OfficeImageFrameTransform? frameTransform) {
        double x = offsetX + drawingShape.X;
        double y = offsetY + drawingShape.Y;
        OfficeShape shape = drawingShape.Shape.Clone();
        if (frameTransform.HasValue && frameTransform.Value.HasTransform) {
            OfficeTransform frame = CreateLocalFrameTransform(frameTransform.Value, x, y);
            shape.Transform = shape.Transform.HasValue ? shape.Transform.Value.Then(frame) : frame;
        }

        AddShape(shape, x, y);
    }

    private void AddNestedText(OfficeDrawingText text, double offsetX, double offsetY, OfficeImageFrameTransform? frameTransform) {
        double x = offsetX + text.X;
        double y = offsetY + text.Y;
        double rotationDegrees = text.RotationDegrees;
        double rotationCenterX = text.RotationCenterX + offsetX;
        double rotationCenterY = text.RotationCenterY + offsetY;
        bool flipHorizontal = text.FlipHorizontal;
        bool flipVertical = text.FlipVertical;
        if (frameTransform.HasValue && frameTransform.Value.HasTransform) {
            OfficeImageFrameTransform frame = frameTransform.Value;
            rotationDegrees += frame.RotationDegrees;
            rotationCenterX = frame.CenterX;
            rotationCenterY = frame.CenterY;
            flipHorizontal ^= frame.FlipHorizontal;
            flipVertical ^= frame.FlipVertical;
        }

        AddText(
            text.Text,
            x,
            y,
            text.Width,
            text.Height,
            text.Font,
            text.Color,
            text.Alignment,
            text.LineHeight,
            text.VerticalAlignment,
            rotationDegrees,
            rotationCenterX,
            rotationCenterY,
            text.WrapText,
            text.ShrinkToFit,
            text.StackedText,
            flipHorizontal,
            flipVertical,
            text.Padding,
            text.ParagraphIndent);
    }

    private void AddNestedRichText(OfficeDrawingRichText richText, double offsetX, double offsetY, OfficeImageFrameTransform? frameTransform) {
        double x = offsetX + richText.X;
        double y = offsetY + richText.Y;
        double rotationDegrees = richText.RotationDegrees;
        double rotationCenterX = richText.RotationCenterX + offsetX;
        double rotationCenterY = richText.RotationCenterY + offsetY;
        bool flipHorizontal = richText.FlipHorizontal;
        bool flipVertical = richText.FlipVertical;
        if (frameTransform.HasValue && frameTransform.Value.HasTransform) {
            OfficeImageFrameTransform frame = frameTransform.Value;
            rotationDegrees += frame.RotationDegrees;
            rotationCenterX = frame.CenterX;
            rotationCenterY = frame.CenterY;
            flipHorizontal ^= frame.FlipHorizontal;
            flipVertical ^= frame.FlipVertical;
        }

        AddRichText(
            richText.Runs,
            x,
            y,
            richText.Width,
            richText.Height,
            richText.Alignment,
            richText.LineHeight,
            richText.VerticalAlignment,
            rotationDegrees,
            rotationCenterX,
            rotationCenterY,
            richText.WrapText,
            richText.ShrinkToFit,
            flipHorizontal,
            flipVertical,
            richText.Padding,
            richText.ParagraphIndent);
    }

    private void AddNestedImage(OfficeDrawingImage image, double offsetX, double offsetY, OfficeImageFrameTransform? frameTransform) {
        OfficeImageProjection projection = image.Projection.Translate(offsetX, offsetY);
        if (frameTransform.HasValue && frameTransform.Value.HasTransform) {
            OfficeImageFrameTransform frame = frameTransform.Value;
            projection = new OfficeImageProjection(
                new OfficeImagePlacement(projection.X, projection.Y, projection.Width, projection.Height),
                projection.SourceCrop,
                projection.RotationDegrees + frame.RotationDegrees,
                frame.CenterX,
                frame.CenterY,
                projection.FlipHorizontal ^ frame.FlipHorizontal,
                projection.FlipVertical ^ frame.FlipVertical);
        }

        AddImage(image.Bytes, image.ContentType, projection, image.AlternativeText);
    }

    private void AddNestedGroup(OfficeDrawingGroup group, double offsetX, double offsetY, OfficeImageFrameTransform? frameTransform) {
        OfficeImageFrameTransform? groupTransform = group.FrameTransform;
        if ((!groupTransform.HasValue || !groupTransform.Value.HasTransform) && frameTransform.HasValue && frameTransform.Value.HasTransform) {
            groupTransform = frameTransform;
        }

        if (groupTransform.HasValue) {
            AddClippedDrawing(group.InnerDrawing, offsetX + group.X, offsetY + group.Y, group.ClipPath, groupTransform.Value);
        } else {
            AddClippedDrawing(group.InnerDrawing, offsetX + group.X, offsetY + group.Y, group.ClipPath);
        }
    }

    private static OfficeTransform CreateLocalFrameTransform(OfficeImageFrameTransform frameTransform, double elementX, double elementY) {
        return OfficeTransform.Translate(elementX, elementY)
            .Then(frameTransform.CreateDestinationTransform())
            .Then(OfficeTransform.Translate(-elementX, -elementY));
    }

    /// <summary>Creates a detached copy of this drawing and all positioned elements.</summary>
    public OfficeDrawing Clone() {
        var clone = new OfficeDrawing(Width, Height);
        for (int i = 0; i < _elements.Count; i++) {
            OfficeDrawingElement element = _elements[i].CloneElement();
            clone._elements.Add(element);
            if (element is OfficeDrawingShape shape) {
                clone._shapes.Add(shape);
            } else if (element is OfficeDrawingImage image) {
                clone._images.Add(image);
            }
        }

        return clone;
    }

    private static void ValidatePositiveFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value <= 0) {
            throw new ArgumentOutOfRangeException(paramName, "Drawing dimensions must be finite positive numbers.");
        }
    }

    private static void ValidateFiniteNonNegative(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value) || value < 0) {
            throw new ArgumentOutOfRangeException(paramName, "Drawing coordinates must be finite non-negative numbers.");
        }
    }
}
