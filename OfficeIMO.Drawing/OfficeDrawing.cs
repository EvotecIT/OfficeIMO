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
    private readonly List<OfficeDrawingImagePattern> _imagePatterns = new List<OfficeDrawingImagePattern>();
    private readonly ReadOnlyCollection<OfficeDrawingImagePattern> _imagePatternsView;
    private readonly List<OfficeDrawingElement> _elements = new List<OfficeDrawingElement>();
    private readonly ReadOnlyCollection<OfficeDrawingElement> _elementsView;
    private readonly HashSet<OfficeDrawingElement> _behindContentElements = new HashSet<OfficeDrawingElement>();

    /// <summary>Drawing width in the caller's layout unit.</summary>
    public double Width { get; }

    /// <summary>Drawing height in the caller's layout unit.</summary>
    public double Height { get; }

    /// <summary>Positioned shapes in paint order.</summary>
    public IReadOnlyList<OfficeDrawingShape> Shapes => _shapesView;

    /// <summary>Positioned images in paint order.</summary>
    public IReadOnlyList<OfficeDrawingImage> Images => _imagesView;

    /// <summary>Clipped image patterns in paint order.</summary>
    public IReadOnlyList<OfficeDrawingImagePattern> ImagePatterns => _imagePatternsView;

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
        _imagePatternsView = new ReadOnlyCollection<OfficeDrawingImagePattern>(_imagePatterns);
        _elementsView = new ReadOnlyCollection<OfficeDrawingElement>(_elements);
    }

    /// <summary>Adds a shape at a local top-left coordinate and returns this drawing.</summary>
    public OfficeDrawing AddShape(OfficeShape shape, double x, double y) {
        var item = new OfficeDrawingShape(shape, x, y);
        if (item.X < 0D || item.Y < 0D || item.X + item.Shape.Width > Width || item.Y + item.Shape.Height > Height) {
            throw new ArgumentOutOfRangeException(nameof(shape), "Drawing shapes must fit inside the drawing bounds.");
        }

        _shapes.Add(item);
        _elements.Add(item);
        return this;
    }

    internal OfficeDrawing AddShapeForClippedRendering(OfficeShape shape, double x, double y) {
        var item = new OfficeDrawingShape(shape, x, y);
        _shapes.Add(item);
        _elements.Add(item);
        return this;
    }

    /// <summary>Adds a shape behind existing foreground content while keeping an initial page background underneath it.</summary>
    public OfficeDrawing AddShapeBehindContent(OfficeShape shape, double x, double y) {
        var item = new OfficeDrawingShape(shape, x, y);
        if (item.X < 0D || item.Y < 0D || item.X + item.Shape.Width > Width || item.Y + item.Shape.Height > Height) {
            throw new ArgumentOutOfRangeException(nameof(shape), "Drawing shapes must fit inside the drawing bounds.");
        }

        int elementIndex = AddBehindContentElement(item);
        _shapes.Insert(GetTypedElementInsertIndex<OfficeDrawingShape>(elementIndex), item);
        return this;
    }

    /// <summary>Adds text inside a local drawing rectangle and returns this drawing.</summary>
    public OfficeDrawing AddText(string text, double x, double y, double width, double height, OfficeFontInfo? font = null, OfficeColor? color = null, OfficeTextAlignment alignment = OfficeTextAlignment.Left, double? lineHeight = null, OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top, double rotationDegrees = 0D, double? rotationCenterX = null, double? rotationCenterY = null, bool wrapText = false, bool shrinkToFit = false, bool stackedText = false, bool flipHorizontal = false, bool flipVertical = false, OfficeTextPadding? padding = null, OfficeTextParagraphIndent? paragraphIndent = null) {
        return AddTextCore(text, x, y, width, height, font, color, alignment, lineHeight, verticalAlignment, rotationDegrees, rotationCenterX, rotationCenterY, wrapText, shrinkToFit, stackedText, flipHorizontal, flipVertical, padding, paragraphIndent, OfficeTextOverflowBehavior.Ellipsis, null, allowOverflow: false);
    }

    /// <summary>
    /// Adds an already-positioned single text run. The frame width may be clipped independently
    /// from the resolved glyph advance retained for backend-consistent positioning.
    /// </summary>
    /// <param name="text">Text content to draw.</param>
    /// <param name="x">Horizontal frame position in drawing units.</param>
    /// <param name="y">Vertical frame position in drawing units.</param>
    /// <param name="width">Clipping frame width in drawing units.</param>
    /// <param name="height">Clipping frame height in drawing units.</param>
    /// <param name="font">Optional font descriptor.</param>
    /// <param name="color">Optional text color.</param>
    /// <param name="alignment">Horizontal alignment inside the frame.</param>
    /// <param name="lineHeight">Optional resolved line height.</param>
    /// <param name="textAdvanceWidth">Resolved horizontal glyph advance, or <see langword="null"/> to use <paramref name="width"/>.</param>
    /// <returns>The current drawing.</returns>
    public OfficeDrawing AddPositionedText(string text, double x, double y, double width, double height, OfficeFontInfo? font = null, OfficeColor? color = null, OfficeTextAlignment alignment = OfficeTextAlignment.Left, double? lineHeight = null, double? textAdvanceWidth = null) =>
        AddTextCore(text, x, y, width, height, font, color, alignment, lineHeight, OfficeTextVerticalAlignment.Top, 0D, null, null, false, false, false, false, false, null, null, OfficeTextOverflowBehavior.Clip, textAdvanceWidth ?? width, allowOverflow: false);

    private OfficeDrawing AddTextCore(string text, double x, double y, double width, double height, OfficeFontInfo? font, OfficeColor? color, OfficeTextAlignment alignment, double? lineHeight, OfficeTextVerticalAlignment verticalAlignment, double rotationDegrees, double? rotationCenterX, double? rotationCenterY, bool wrapText, bool shrinkToFit, bool stackedText, bool flipHorizontal, bool flipVertical, OfficeTextPadding? padding, OfficeTextParagraphIndent? paragraphIndent, OfficeTextOverflowBehavior overflowBehavior, double? textAdvanceWidth, bool allowOverflow) {
        var item = new OfficeDrawingText(text, x, y, width, height, font, color, alignment, lineHeight, verticalAlignment, rotationDegrees, rotationCenterX, rotationCenterY, wrapText, shrinkToFit, stackedText, flipHorizontal, flipVertical, padding, paragraphIndent, overflowBehavior, textAdvanceWidth);
        if (!allowOverflow && (item.X < 0D || item.Y < 0D || item.X + item.Width > Width || item.Y + item.Height > Height)) {
            throw new ArgumentOutOfRangeException(nameof(text), "Drawing text must fit inside the drawing bounds.");
        }

        _elements.Add(item);
        return this;
    }

    /// <summary>Adds text behind existing foreground content while keeping an initial page background underneath it.</summary>
    public OfficeDrawing AddTextBehindContent(string text, double x, double y, double width, double height, OfficeFontInfo? font = null, OfficeColor? color = null, OfficeTextAlignment alignment = OfficeTextAlignment.Left, double? lineHeight = null, OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top, double rotationDegrees = 0D, double? rotationCenterX = null, double? rotationCenterY = null, bool wrapText = false, bool shrinkToFit = false, bool stackedText = false, bool flipHorizontal = false, bool flipVertical = false, OfficeTextPadding? padding = null, OfficeTextParagraphIndent? paragraphIndent = null) {
        var item = new OfficeDrawingText(text, x, y, width, height, font, color, alignment, lineHeight, verticalAlignment, rotationDegrees, rotationCenterX, rotationCenterY, wrapText, shrinkToFit, stackedText, flipHorizontal, flipVertical, padding, paragraphIndent);
        if (item.X < 0D || item.Y < 0D || item.X + item.Width > Width || item.Y + item.Height > Height) {
            throw new ArgumentOutOfRangeException(nameof(text), "Drawing text must fit inside the drawing bounds.");
        }

        AddBehindContentElement(item);
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

    /// <summary>Adds rich text behind existing foreground content while keeping an initial page background underneath it.</summary>
    public OfficeDrawing AddRichTextBehindContent(IReadOnlyList<OfficeRichTextRun> runs, double x, double y, double width, double height, OfficeTextAlignment alignment = OfficeTextAlignment.Left, double? lineHeight = null, OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top, double rotationDegrees = 0D, double? rotationCenterX = null, double? rotationCenterY = null, bool wrapText = true, bool shrinkToFit = false, bool flipHorizontal = false, bool flipVertical = false, OfficeTextPadding? padding = null, OfficeTextParagraphIndent? paragraphIndent = null) {
        var item = new OfficeDrawingRichText(runs, x, y, width, height, alignment, lineHeight, verticalAlignment, rotationDegrees, rotationCenterX, rotationCenterY, wrapText, shrinkToFit, flipHorizontal, flipVertical, padding, paragraphIndent);
        if (item.X + item.Width > Width || item.Y + item.Height > Height) {
            throw new ArgumentOutOfRangeException(nameof(runs), "Drawing rich text must fit inside the drawing bounds.");
        }

        AddBehindContentElement(item);
        return this;
    }

    /// <summary>Adds an image using a shared placement/crop/transform projection and returns this drawing.</summary>
    public OfficeDrawing AddImage(byte[] bytes, string? contentType, OfficeImageProjection projection, string? alternativeText = null, double opacity = 1D) {
        return AddImageCore(bytes, contentType, projection, alternativeText, opacity, allowOverflow: false);
    }

    internal OfficeDrawing AddImageShared(byte[] bytes, string? contentType, OfficeImageProjection projection,
        string? alternativeText = null, double opacity = 1D) {
        return AddImageCore(bytes, contentType, projection, alternativeText, opacity,
            allowOverflow: false, useDataSnapshot: true);
    }

    /// <summary>Adds an image clipped by a drawing-local clipping path.</summary>
    public OfficeDrawing AddClippedImage(byte[] bytes, string? contentType, OfficeImageProjection projection, double clipX, double clipY, OfficeClipPath clipPath, string? alternativeText = null, double opacity = 1D) {
        if (clipPath == null) {
            throw new ArgumentNullException(nameof(clipPath));
        }

        ValidateFiniteNonNegative(clipX, nameof(clipX));
        ValidateFiniteNonNegative(clipY, nameof(clipY));
        if (clipX + clipPath.Width > Width || clipY + clipPath.Height > Height) {
            throw new ArgumentOutOfRangeException(nameof(clipPath), "Image clip must fit inside the drawing bounds.");
        }

        var clipped = new OfficeDrawing(clipPath.Width, clipPath.Height);
        clipped.AddImageCore(bytes, contentType, projection.Translate(-clipX, -clipY), alternativeText, opacity, allowOverflow: true);
        return AddClippedDrawing(clipped, clipX, clipY, clipPath);
    }

    /// <summary>Adds text clipped by a drawing-local clipping path.</summary>
    public OfficeDrawing AddClippedText(string text, double x, double y, double width, double height, double clipX, double clipY, OfficeClipPath clipPath, OfficeFontInfo? font = null, OfficeColor? color = null, OfficeTextAlignment alignment = OfficeTextAlignment.Left, double? lineHeight = null, OfficeTextVerticalAlignment verticalAlignment = OfficeTextVerticalAlignment.Top, double rotationDegrees = 0D, double? rotationCenterX = null, double? rotationCenterY = null, bool wrapText = false, bool shrinkToFit = false, bool stackedText = false, bool flipHorizontal = false, bool flipVertical = false, OfficeTextPadding? padding = null, OfficeTextParagraphIndent? paragraphIndent = null) {
        if (clipPath == null) {
            throw new ArgumentNullException(nameof(clipPath));
        }

        ValidateFiniteNonNegative(clipX, nameof(clipX));
        ValidateFiniteNonNegative(clipY, nameof(clipY));
        if (clipX + clipPath.Width > Width || clipY + clipPath.Height > Height) {
            throw new ArgumentOutOfRangeException(nameof(clipPath), "Text clip must fit inside the drawing bounds.");
        }

        var clipped = new OfficeDrawing(clipPath.Width, clipPath.Height);
        clipped.AddTextCore(
            text,
            x - clipX,
            y - clipY,
            width,
            height,
            font,
            color,
            alignment,
            lineHeight,
            verticalAlignment,
            rotationDegrees,
            rotationCenterX.HasValue ? rotationCenterX.Value - clipX : null,
            rotationCenterY.HasValue ? rotationCenterY.Value - clipY : null,
            wrapText,
            shrinkToFit,
            stackedText,
            flipHorizontal,
            flipVertical,
            padding,
            paragraphIndent,
            OfficeTextOverflowBehavior.Ellipsis,
            null,
            allowOverflow: true);
        return AddClippedDrawing(clipped, clipX, clipY, clipPath);
    }

    private OfficeDrawing AddImageCore(byte[] bytes, string? contentType, OfficeImageProjection projection,
        string? alternativeText, double opacity, bool allowOverflow, bool useDataSnapshot = false) {
        var item = new OfficeDrawingImage(bytes, contentType, projection, alternativeText, opacity,
            useDataSnapshot);
        (double left, double top, double right, double bottom) = item.Projection.GetDestinationBounds();
        if (!allowOverflow && (left < 0D || top < 0D || right > Width || bottom > Height)) {
            throw new ArgumentOutOfRangeException(nameof(projection), "Drawing images must fit inside the drawing bounds.");
        }

        _images.Add(item);
        _elements.Add(item);
        return this;
    }

    /// <summary>Adds an image behind existing foreground content while keeping an initial page background underneath it.</summary>
    public OfficeDrawing AddImageBehindContent(byte[] bytes, string? contentType, OfficeImageProjection projection, string? alternativeText = null, double opacity = 1D) {
        return AddImageBehindContentCore(bytes, contentType, projection, alternativeText, opacity, allowOverflow: false);
    }

    private OfficeDrawing AddImageBehindContentCore(byte[] bytes, string? contentType, OfficeImageProjection projection, string? alternativeText, double opacity, bool allowOverflow) {
        var item = new OfficeDrawingImage(bytes, contentType, projection, alternativeText, opacity);
        (double left, double top, double right, double bottom) = item.Projection.GetDestinationBounds();
        if (!allowOverflow && (left < 0D || top < 0D || right > Width || bottom > Height)) {
            throw new ArgumentOutOfRangeException(nameof(projection), "Drawing images must fit inside the drawing bounds.");
        }

        int elementIndex = AddBehindContentElement(item);
        _images.Insert(GetTypedElementInsertIndex<OfficeDrawingImage>(elementIndex), item);
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

    /// <summary>Adds another drawing as one affine-transformed, isolated opacity group.</summary>
    public OfficeDrawing AddEffectDrawing(OfficeDrawing drawing, OfficeTransform transform, double opacity = 1D) {
        if (drawing == null) throw new ArgumentNullException(nameof(drawing));
        Fonts.AddRange(drawing.Fonts);
        _elements.Add(new OfficeDrawingEffectGroup(drawing, transform, opacity));
        return this;
    }

    /// <summary>Adds another drawing as one affine-transformed group with managed blending and an optional vector soft mask.</summary>
    public OfficeDrawing AddEffectDrawing(
        OfficeDrawing drawing,
        OfficeTransform transform,
        OfficeBlendMode blendMode,
        OfficeDrawingSoftMask? softMask = null,
        double opacity = 1D) {
        if (drawing == null) throw new ArgumentNullException(nameof(drawing));
        Fonts.AddRange(drawing.Fonts);
        if (softMask != null) Fonts.AddRange(softMask.InnerDrawing.Fonts);
        _elements.Add(new OfficeDrawingEffectGroup(drawing, transform, blendMode, softMask, opacity));
        return this;
    }

    /// <summary>Adds another drawing as a clipped nested group at a local destination offset.</summary>
    public OfficeDrawing AddClippedDrawing(OfficeDrawing drawing, double x, double y, OfficeClipPath clipPath) {
        return AddClippedDrawingCore(drawing, x, y, clipPath, 0D, 0D, null);
    }

    /// <summary>Adds another drawing as a clipped nested group at a local destination offset with a shared frame transform.</summary>
    public OfficeDrawing AddClippedDrawing(OfficeDrawing drawing, double x, double y, OfficeClipPath clipPath, OfficeImageFrameTransform frameTransform) {
        return AddClippedDrawingCore(drawing, x, y, clipPath, 0D, 0D, frameTransform);
    }

    /// <summary>Adds another drawing as a clipped group with an independent content offset.</summary>
    public OfficeDrawing AddClippedDrawing(OfficeDrawing drawing, double x, double y, OfficeClipPath clipPath, double contentOffsetX, double contentOffsetY) {
        return AddClippedDrawingCore(drawing, x, y, clipPath, contentOffsetX, contentOffsetY, null);
    }

    /// <summary>Adds another drawing as a clipped group with independent content offset and frame transform.</summary>
    public OfficeDrawing AddClippedDrawing(OfficeDrawing drawing, double x, double y, OfficeClipPath clipPath, double contentOffsetX, double contentOffsetY, OfficeImageFrameTransform frameTransform) {
        return AddClippedDrawingCore(drawing, x, y, clipPath, contentOffsetX, contentOffsetY, frameTransform);
    }

    private OfficeDrawing AddClippedDrawingCore(OfficeDrawing drawing, double x, double y, OfficeClipPath clipPath, double contentOffsetX, double contentOffsetY, OfficeImageFrameTransform? frameTransform) {
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

        Fonts.AddRange(drawing.Fonts);
        _elements.Add(new OfficeDrawingGroup(drawing, x, y, clipPath, contentOffsetX, contentOffsetY, frameTransform));
        return this;
    }

    private OfficeDrawing AddDrawingCore(OfficeDrawing drawing, double x, double y, OfficeImageFrameTransform? frameTransform) {
        return AddDrawingCore(drawing, x, y, frameTransform, allowOverflow: false);
    }

    internal OfficeDrawing AddDrawingForClippedRendering(OfficeDrawing drawing, double x, double y, OfficeImageFrameTransform? frameTransform) {
        return AddDrawingCore(drawing, x, y, frameTransform, allowOverflow: true);
    }

    private OfficeDrawing AddDrawingCore(OfficeDrawing drawing, double x, double y, OfficeImageFrameTransform? frameTransform, bool allowOverflow) {
        if (drawing == null) {
            throw new ArgumentNullException(nameof(drawing));
        }

        if (allowOverflow) {
            ValidateFinite(x, nameof(x));
            ValidateFinite(y, nameof(y));
        } else {
            ValidateFiniteNonNegative(x, nameof(x));
            ValidateFiniteNonNegative(y, nameof(y));
        }
        if (!allowOverflow && (x + drawing.Width > Width || y + drawing.Height > Height)) {
            throw new ArgumentOutOfRangeException(nameof(drawing), "Nested drawing content must fit inside the drawing bounds.");
        }

        if (!allowOverflow && frameTransform.HasValue && frameTransform.Value.HasTransform && ContainsImagePattern(drawing)) {
            AddNestedGroupElement(drawing, x, y, OfficeClipPath.Rectangle(drawing.Width, drawing.Height), 0D, 0D, frameTransform, allowOverflow);
            return this;
        }

        Fonts.AddRange(drawing.Fonts);
        for (int i = 0; i < drawing.Elements.Count; i++) {
            OfficeDrawingElement element = drawing.Elements[i];
            if (element is OfficeDrawingShape shape) {
                AddNestedShape(shape, x, y, frameTransform, allowOverflow);
            } else if (element is OfficeDrawingText text) {
                AddNestedText(text, x, y, frameTransform, allowOverflow);
            } else if (element is OfficeDrawingRichText richText) {
                AddNestedRichText(richText, x, y, frameTransform, allowOverflow);
            } else if (element is OfficeDrawingImage image) {
                AddNestedImage(image, x, y, frameTransform, allowOverflow);
            } else if (element is OfficeDrawingImagePattern imagePattern) {
                AddNestedImagePattern(imagePattern, x, y, frameTransform, allowOverflow);
            } else if (element is OfficeDrawingTilingPattern tilingPattern) {
                AddNestedTilingPattern(tilingPattern, x, y, frameTransform, allowOverflow);
            } else if (element is OfficeDrawingEffectGroup effectGroup) {
                OfficeTransform translatedTransform = effectGroup.Transform.Then(OfficeTransform.Translate(x, y));
                if (frameTransform.HasValue && frameTransform.Value.HasTransform) {
                    translatedTransform = translatedTransform.Then(frameTransform.Value.CreateDestinationTransform());
                }
                AddEffectDrawing(effectGroup.InnerDrawing, translatedTransform, effectGroup.BlendMode, effectGroup.SoftMask, effectGroup.Opacity);
            } else if (element is OfficeDrawingGroup group) {
                AddNestedGroup(group, x, y, frameTransform, allowOverflow);
            }
        }

        return this;
    }

    private void AddNestedShape(OfficeDrawingShape drawingShape, double offsetX, double offsetY, OfficeImageFrameTransform? frameTransform, bool allowOverflow) {
        double x = offsetX + drawingShape.X;
        double y = offsetY + drawingShape.Y;
        OfficeShape shape = drawingShape.Shape.Clone();
        if (frameTransform.HasValue && frameTransform.Value.HasTransform) {
            OfficeTransform frame = CreateLocalFrameTransform(frameTransform.Value, x, y);
            shape.Transform = shape.Transform.HasValue ? shape.Transform.Value.Then(frame) : frame;
        }

        if (allowOverflow) {
            var item = new OfficeDrawingShape(shape, x, y);
            _shapes.Add(item);
            _elements.Add(item);
        } else {
            AddShape(shape, x, y);
        }
    }

    private void AddNestedText(OfficeDrawingText text, double offsetX, double offsetY, OfficeImageFrameTransform? frameTransform, bool allowOverflow) {
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

        var item = new OfficeDrawingText(
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
            text.ParagraphIndent,
            text.OverflowBehavior,
            text.TextAdvanceWidth);
        if (!allowOverflow && (item.X + item.Width > Width || item.Y + item.Height > Height)) {
            throw new ArgumentOutOfRangeException(nameof(text), "Drawing text must fit inside the drawing bounds.");
        }

        _elements.Add(item);
    }

    private void AddNestedRichText(OfficeDrawingRichText richText, double offsetX, double offsetY, OfficeImageFrameTransform? frameTransform, bool allowOverflow) {
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

        var item = new OfficeDrawingRichText(
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
        if (!allowOverflow && (item.X + item.Width > Width || item.Y + item.Height > Height)) {
            throw new ArgumentOutOfRangeException(nameof(richText), "Drawing rich text must fit inside the drawing bounds.");
        }

        _elements.Add(item);
    }

    private void AddNestedImage(OfficeDrawingImage image, double offsetX, double offsetY, OfficeImageFrameTransform? frameTransform, bool allowOverflow) {
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

        if (allowOverflow) {
            var item = new OfficeDrawingImage(image.EncodedBytes, image.ContentType, projection, image.AlternativeText, image.Opacity, useDataSnapshot: true);
            _images.Add(item);
            _elements.Add(item);
        } else {
            AddImage(image.EncodedBytes, image.ContentType, projection, image.AlternativeText, image.Opacity);
        }
    }

    private void AddNestedGroup(OfficeDrawingGroup group, double offsetX, double offsetY, OfficeImageFrameTransform? frameTransform, bool allowOverflow) {
        OfficeImageFrameTransform? groupTransform = group.FrameTransform;
        if (groupTransform.HasValue && groupTransform.Value.HasTransform && frameTransform.HasValue && frameTransform.Value.HasTransform) {
            double wrapperWidth = group.X + group.ClipPath.Width;
            double wrapperHeight = group.Y + group.ClipPath.Height;
            var wrapper = new OfficeDrawing(wrapperWidth, wrapperHeight);
            wrapper.AddClippedDrawing(group.InnerDrawing, group.X, group.Y, group.ClipPath, group.ContentOffsetX, group.ContentOffsetY, groupTransform.Value);
            AddNestedGroupElement(wrapper, offsetX, offsetY, OfficeClipPath.Rectangle(wrapperWidth, wrapperHeight), 0D, 0D, frameTransform.Value, allowOverflow);
            return;
        }

        if ((!groupTransform.HasValue || !groupTransform.Value.HasTransform) && frameTransform.HasValue && frameTransform.Value.HasTransform) {
            groupTransform = frameTransform;
        }

        if (groupTransform.HasValue) {
            AddNestedGroupElement(group.InnerDrawing, offsetX + group.X, offsetY + group.Y, group.ClipPath, group.ContentOffsetX, group.ContentOffsetY, groupTransform.Value, allowOverflow);
        } else {
            AddNestedGroupElement(group.InnerDrawing, offsetX + group.X, offsetY + group.Y, group.ClipPath, group.ContentOffsetX, group.ContentOffsetY, null, allowOverflow);
        }
    }

    private void AddNestedGroupElement(OfficeDrawing drawing, double x, double y, OfficeClipPath clipPath, double contentOffsetX, double contentOffsetY, OfficeImageFrameTransform? frameTransform, bool allowOverflow) {
        if (allowOverflow) {
            _elements.Add(new OfficeDrawingGroup(drawing, x, y, clipPath, contentOffsetX, contentOffsetY, frameTransform));
        } else if (frameTransform.HasValue) {
            AddClippedDrawing(drawing, x, y, clipPath, contentOffsetX, contentOffsetY, frameTransform.Value);
        } else {
            AddClippedDrawing(drawing, x, y, clipPath, contentOffsetX, contentOffsetY);
        }
    }

    private static OfficeTransform CreateLocalFrameTransform(OfficeImageFrameTransform frameTransform, double elementX, double elementY) {
        return OfficeTransform.Translate(elementX, elementY)
            .Then(frameTransform.CreateDestinationTransform())
            .Then(OfficeTransform.Translate(-elementX, -elementY));
    }

    private int GetBehindContentInsertIndex() {
        if (_elements.Count == 0) {
            return 0;
        }

        int index = 0;
        if (_elements[0] is OfficeDrawingShape shape &&
            shape.X == 0D &&
            shape.Y == 0D &&
            shape.Shape.Kind == OfficeShapeKind.Rectangle &&
            shape.Shape.Width == Width &&
            shape.Shape.Height == Height &&
            shape.Shape.StrokeWidth <= 0D) {
            index = 1;
        }

        while (index < _elements.Count && _behindContentElements.Contains(_elements[index])) {
            index++;
        }

        return index;
    }

    private int AddBehindContentElement(OfficeDrawingElement item) {
        _behindContentElements.Add(item);
        int index = GetBehindContentInsertIndex();
        _elements.Insert(index, item);
        return index;
    }

    private int GetTypedElementInsertIndex<T>(int elementIndex) where T : OfficeDrawingElement {
        int index = 0;
        for (int i = 0; i < elementIndex; i++) {
            if (_elements[i] is T) {
                index++;
            }
        }

        return index;
    }

    /// <summary>Creates a detached copy of this drawing and all positioned elements.</summary>
    public OfficeDrawing Clone() {
        var clone = new OfficeDrawing(Width, Height);
        for (int i = 0; i < _elements.Count; i++) {
            OfficeDrawingElement element = _elements[i].CloneElement();
            clone._elements.Add(element);
            if (_behindContentElements.Contains(_elements[i])) {
                clone._behindContentElements.Add(element);
            }

            if (element is OfficeDrawingShape shape) {
                clone._shapes.Add(shape);
            } else if (element is OfficeDrawingImage image) {
                clone._images.Add(image);
            } else if (element is OfficeDrawingImagePattern imagePattern) {
                clone._imagePatterns.Add(imagePattern);
            }
        }

        return clone;
    }

    private static bool ContainsImagePattern(OfficeDrawing drawing) {
        for (int index = 0; index < drawing.Elements.Count; index++) {
            if (drawing.Elements[index] is OfficeDrawingImagePattern) return true;
            if (drawing.Elements[index] is OfficeDrawingGroup group && ContainsImagePattern(group.InnerDrawing)) return true;
            if (drawing.Elements[index] is OfficeDrawingEffectGroup effectGroup && ContainsImagePattern(effectGroup.InnerDrawing)) return true;
        }

        return false;
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

    private static void ValidateFinite(double value, string paramName) {
        if (double.IsNaN(value) || double.IsInfinity(value)) {
            throw new ArgumentOutOfRangeException(paramName, "Drawing coordinates must be finite numbers.");
        }
    }
}
