using OfficeIMO.Drawing;

namespace OfficeIMO.Pdf;

public sealed partial class PdfDocument {
    /// <summary>Adds foreground page content at absolute top-left page coordinates.</summary>
    public PdfDocument Canvas(System.Action<PdfPageCanvas> build) {
        Guard.NotNull(build, nameof(build));
        var canvas = new PdfPageCanvas();
        build(canvas);
        AddBlock(new PdfCanvasBlock(canvas.Items));
        return this;
    }

    /// <summary>Adds a shared OfficeIMO.Drawing shape at the current flow position.</summary>
    public PdfDocument Shape(OfficeShape shape, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        AddBlock(CreateShapeBlock(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents));
        return this;
    }

    /// <summary>Adds a shared OfficeIMO.Drawing scene at the current flow position.</summary>
    public PdfDocument Drawing(OfficeDrawing drawing, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        AddBlock(CreateDrawingBlock(drawing, align, spacingBefore, spacingAfter, style, linkUri, linkContents));
        return this;
    }

    /// <summary>Adds a flow line using the shared OfficeIMO.Drawing shape descriptor.</summary>
    public PdfDocument Line(double x1, double y1, double x2, double y2, PdfColor? strokeColor = null, double strokeWidth = 1, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        var shape = OfficeShape.Line(x1, y1, x2, y2);
        shape.StrokeColor = (strokeColor ?? PdfColor.Gray).ToOfficeColor();
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = strokeDashStyle;
        shape.StrokeLineCap = strokeLineCap;
        shape.StrokeLineJoin = strokeLineJoin;
        return Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
    }

    /// <summary>Adds a flow rectangle using the shared OfficeIMO.Drawing shape descriptor.</summary>
    public PdfDocument Rectangle(double width, double height, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        var shape = OfficeShape.Rectangle(width, height);
        shape.StrokeColor = (strokeColor ?? PdfColor.Gray).ToOfficeColor();
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = strokeDashStyle;
        shape.StrokeLineCap = strokeLineCap;
        shape.StrokeLineJoin = strokeLineJoin;
        shape.FillColor = fillColor?.ToOfficeColor();
        return Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
    }

    /// <summary>Adds a flow rounded rectangle using the shared OfficeIMO.Drawing shape descriptor.</summary>
    public PdfDocument RoundedRectangle(double width, double height, double cornerRadius, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        var shape = OfficeShape.RoundedRectangle(width, height, cornerRadius);
        shape.StrokeColor = (strokeColor ?? PdfColor.Gray).ToOfficeColor();
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = strokeDashStyle;
        shape.StrokeLineCap = strokeLineCap;
        shape.StrokeLineJoin = strokeLineJoin;
        shape.FillColor = fillColor?.ToOfficeColor();
        return Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
    }

    /// <summary>Adds a flow ellipse using the shared OfficeIMO.Drawing shape descriptor.</summary>
    public PdfDocument Ellipse(double width, double height, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        var shape = OfficeShape.Ellipse(width, height);
        shape.StrokeColor = (strokeColor ?? PdfColor.Gray).ToOfficeColor();
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = strokeDashStyle;
        shape.StrokeLineCap = strokeLineCap;
        shape.StrokeLineJoin = strokeLineJoin;
        shape.FillColor = fillColor?.ToOfficeColor();
        return Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
    }

    /// <summary>Adds a flow polygon using the shared OfficeIMO.Drawing shape descriptor.</summary>
    public PdfDocument Polygon(System.Collections.Generic.IEnumerable<OfficePoint> points, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        var shape = OfficeShape.Polygon(points);
        shape.StrokeColor = (strokeColor ?? PdfColor.Gray).ToOfficeColor();
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = strokeDashStyle;
        shape.StrokeLineCap = strokeLineCap;
        shape.StrokeLineJoin = strokeLineJoin;
        shape.FillColor = fillColor?.ToOfficeColor();
        return Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
    }

    /// <summary>Adds a flow path using the shared OfficeIMO.Drawing shape descriptor.</summary>
    public PdfDocument Path(System.Collections.Generic.IEnumerable<OfficePathCommand> commands, PdfColor? strokeColor = null, double strokeWidth = 1, PdfColor? fillColor = null, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, OfficeStrokeDashStyle strokeDashStyle = OfficeStrokeDashStyle.Solid, OfficeStrokeLineCap? strokeLineCap = null, OfficeStrokeLineJoin? strokeLineJoin = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        var shape = OfficeShape.Path(commands);
        shape.StrokeColor = (strokeColor ?? PdfColor.Gray).ToOfficeColor();
        shape.StrokeWidth = strokeWidth;
        shape.StrokeDashStyle = strokeDashStyle;
        shape.StrokeLineCap = strokeLineCap;
        shape.StrokeLineJoin = strokeLineJoin;
        shape.FillColor = fillColor?.ToOfficeColor();
        return Shape(shape, align, spacingBefore, spacingAfter, style, linkUri, linkContents);
    }

    /// <summary>Adds a raster image supported by OfficeIMO.Drawing at the current flow position.</summary>
    public PdfDocument Image(byte[] jpegBytes, double width, double height, PdfAlign? align = null, OfficeClipPath? clipPath = null, OfficeImageFit? fit = null, double? spacingBefore = null, double? spacingAfter = null, PdfImageStyle? style = null, string? linkUri = null, string? linkContents = null) =>
        Image(jpegBytes, width, height, align, clipPath, fit, spacingBefore, spacingAfter, style, linkUri, linkContents, alternativeText: null);

    /// <summary>Adds a supported meaningful image at the current flow position with alternate text.</summary>
    public PdfDocument Image(byte[] jpegBytes, double width, double height, string? alternativeText) =>
        Image(jpegBytes, width, height, align: null, clipPath: null, fit: null, spacingBefore: null, spacingAfter: null, style: null, linkUri: null, linkContents: null, alternativeText: alternativeText);

    /// <summary>Adds a raster image supported by OfficeIMO.Drawing at the current flow position.</summary>
    public PdfDocument Image(byte[] jpegBytes, double width, double height, PdfAlign? align, OfficeClipPath? clipPath, OfficeImageFit? fit, double? spacingBefore, double? spacingAfter, PdfImageStyle? style, string? linkUri, string? linkContents, string? alternativeText) {
        Guard.NotNullOrEmpty(jpegBytes, nameof(jpegBytes));
        Guard.Positive(width, nameof(width));
        Guard.Positive(height, nameof(height));
        Guard.OptionalUriAction(linkUri, nameof(linkUri));
        PdfImageStyle? imageStyle = CreateImageStyle(align, clipPath, fit, spacingBefore, spacingAfter, style, alternativeText);
        if (imageStyle != null) {
            ValidateImageStyleForBox(imageStyle, width, height, nameof(clipPath));
        }

        PreparedImage prepared = PrepareImageBytes(jpegBytes);
        if (imageStyle != null) {
            ValidateImageFitDimensions(prepared.Info, imageStyle.Fit, nameof(fit));
        }

        AddBlock(new ImageBlock(prepared.Data, width, height, prepared.Info, imageStyle, linkUri, linkContents, useDataSnapshot: true));
        return this;
    }

    internal static ShapeBlock CreateShapeBlock(OfficeShape shape, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        Guard.NotNull(shape, nameof(shape));
        if (shape.Kind == OfficeShapeKind.Line) {
            Guard.NonNegative(shape.Width, nameof(shape.Width));
            Guard.NonNegative(shape.Height, nameof(shape.Height));
        } else {
            Guard.Positive(shape.Width, nameof(shape.Width));
            Guard.Positive(shape.Height, nameof(shape.Height));
        }
        Guard.NonNegative(shape.StrokeWidth, nameof(shape.StrokeWidth));
        Guard.OptionalUriAction(linkUri, nameof(linkUri));
        ValidateOpacity(shape.FillOpacity, nameof(shape.FillOpacity));
        ValidateOpacity(shape.StrokeOpacity, nameof(shape.StrokeOpacity));
        ValidateShapeClipPath(shape);
        if (shape.Kind != OfficeShapeKind.Line && shape.Kind != OfficeShapeKind.Rectangle && shape.Kind != OfficeShapeKind.RoundedRectangle && shape.Kind != OfficeShapeKind.Ellipse && shape.Kind != OfficeShapeKind.Polygon && shape.Kind != OfficeShapeKind.Path) {
            throw new System.NotSupportedException($"OfficeIMO.Pdf currently supports {nameof(OfficeShapeKind.Line)}, {nameof(OfficeShapeKind.Rectangle)}, {nameof(OfficeShapeKind.RoundedRectangle)}, {nameof(OfficeShapeKind.Ellipse)}, {nameof(OfficeShapeKind.Polygon)}, and {nameof(OfficeShapeKind.Path)} shapes only.");
        }

        if (shape.Kind == OfficeShapeKind.Line) {
            if (shape.Points.Count != 2 || shape.Points[0] == shape.Points[1]) {
                throw new System.ArgumentException("Line shapes require exactly two different points.", nameof(shape));
            }

            for (int i = 0; i < shape.Points.Count; i++) {
                ValidatePointInsideShape(shape.Points[i], shape);
            }
        }

        if (shape.Kind == OfficeShapeKind.RoundedRectangle) {
            Guard.NonNegative(shape.CornerRadius, nameof(shape.CornerRadius));
            if (shape.CornerRadius > System.Math.Min(shape.Width, shape.Height) / 2D) {
                throw new System.ArgumentOutOfRangeException(nameof(shape), "Rounded rectangle corner radius cannot exceed half of the shape width or height.");
            }
        }

        if (shape.Kind == OfficeShapeKind.Polygon) {
            if (shape.Points.Count < 3) {
                throw new System.ArgumentException("Polygon shapes require at least three points.", nameof(shape));
            }

            for (int i = 0; i < shape.Points.Count; i++) {
                var point = shape.Points[i];
                Guard.NonNegative(point.X, nameof(shape.Points));
                Guard.NonNegative(point.Y, nameof(shape.Points));
                if (point.X > shape.Width || point.Y > shape.Height) {
                    throw new System.ArgumentOutOfRangeException(nameof(shape), "Polygon points must fit inside the shape width and height.");
                }
            }
        }

        if (shape.Kind == OfficeShapeKind.Path) {
            if (shape.PathCommands.Count == 0 || shape.PathCommands[0].Kind != OfficePathCommandKind.MoveTo) {
                throw new System.ArgumentException("Path shapes require commands starting with MoveTo.", nameof(shape));
            }

            bool hasDraw = false;
            for (int i = 0; i < shape.PathCommands.Count; i++) {
                var command = shape.PathCommands[i];
                switch (command.Kind) {
                    case OfficePathCommandKind.MoveTo:
                        ValidatePointInsideShape(command.Point, shape);
                        break;
                    case OfficePathCommandKind.LineTo:
                        ValidatePointInsideShape(command.Point, shape);
                        hasDraw = true;
                        break;
                    case OfficePathCommandKind.QuadraticBezierTo:
                        ValidatePointInsideShape(command.ControlPoint1, shape);
                        ValidatePointInsideShape(command.Point, shape);
                        hasDraw = true;
                        break;
                    case OfficePathCommandKind.CubicBezierTo:
                        ValidatePointInsideShape(command.ControlPoint1, shape);
                        ValidatePointInsideShape(command.ControlPoint2, shape);
                        ValidatePointInsideShape(command.Point, shape);
                        hasDraw = true;
                        break;
                    case OfficePathCommandKind.Close:
                        break;
                    default:
                        throw new System.ArgumentOutOfRangeException(nameof(shape), "Unsupported path command kind.");
                }
            }

            if (!hasDraw) {
                throw new System.ArgumentException("Path shapes require at least one drawing command.", nameof(shape));
            }
        }

        PdfDrawingStyle? drawingStyle = CreateDrawingStyle(align, spacingBefore, spacingAfter, style, "Shape");
        if (drawingStyle != null) {
            ValidateDrawingStyle(drawingStyle, "Shape");
        }

        return new ShapeBlock(shape, drawingStyle, linkUri, linkContents);
    }

    internal static PdfHorizontalRuleStyle? CreateHorizontalRuleStyle(double? thickness, PdfColor? color, double? spacingBefore, double? spacingAfter, PdfHorizontalRuleStyle? style) {
        if (!thickness.HasValue && !color.HasValue && !spacingBefore.HasValue && !spacingAfter.HasValue && style == null) {
            return null;
        }

        var ruleStyle = style?.Clone() ?? new PdfHorizontalRuleStyle();
        if (thickness.HasValue) {
            ruleStyle.Thickness = thickness.Value;
        }

        if (color.HasValue) {
            ruleStyle.Color = color.Value;
        }

        if (spacingBefore.HasValue) {
            ruleStyle.SpacingBefore = spacingBefore.Value;
        }

        if (spacingAfter.HasValue) {
            ruleStyle.SpacingAfter = spacingAfter.Value;
        }

        return ruleStyle;
    }

    internal static PdfImageStyle? CreateImageStyle(PdfAlign? align, OfficeClipPath? clipPath, OfficeImageFit? fit, double? spacingBefore, double? spacingAfter, PdfImageStyle? style, string? alternativeText = null) {
        if (!align.HasValue && clipPath == null && !fit.HasValue && !spacingBefore.HasValue && !spacingAfter.HasValue && style == null && alternativeText == null) {
            return null;
        }

        var imageStyle = style?.Clone() ?? new PdfImageStyle();
        if (align.HasValue) {
            imageStyle.Align = align.Value;
        }

        if (clipPath != null) {
            imageStyle.ClipPath = clipPath;
        }

        if (fit.HasValue) {
            ValidateImageFit(fit.Value, nameof(fit));
            imageStyle.Fit = fit.Value;
        }

        if (spacingBefore.HasValue) {
            imageStyle.SpacingBefore = spacingBefore.Value;
        }

        if (spacingAfter.HasValue) {
            imageStyle.SpacingAfter = spacingAfter.Value;
        }

        if (alternativeText != null) {
            imageStyle.AlternativeText = alternativeText;
        }

        return imageStyle;
    }

    internal static PdfDrawingStyle? CreateDrawingStyle(PdfAlign? align, double? spacingBefore, double? spacingAfter, PdfDrawingStyle? style, string objectName = "Drawing") {
        if (!align.HasValue && !spacingBefore.HasValue && !spacingAfter.HasValue && style == null) {
            return null;
        }

        var drawingStyle = style?.Clone() ?? new PdfDrawingStyle();
        if (align.HasValue) {
            Guard.LeftCenterRightAlign(align.Value, nameof(align), objectName);
            drawingStyle.Align = align.Value;
        }

        if (spacingBefore.HasValue) {
            if (spacingBefore.Value < 0 || double.IsNaN(spacingBefore.Value) || double.IsInfinity(spacingBefore.Value)) {
                throw new System.ArgumentException(objectName + " spacing before must be a non-negative finite value.", nameof(spacingBefore));
            }

            drawingStyle.SpacingBefore = spacingBefore.Value;
        }

        if (spacingAfter.HasValue) {
            if (spacingAfter.Value < 0 || double.IsNaN(spacingAfter.Value) || double.IsInfinity(spacingAfter.Value)) {
                throw new System.ArgumentException(objectName + " spacing after must be a non-negative finite value.", nameof(spacingAfter));
            }

            drawingStyle.SpacingAfter = spacingAfter.Value;
        }

        return drawingStyle;
    }

    internal static DrawingBlock CreateDrawingBlock(OfficeDrawing drawing, PdfAlign? align = null, double? spacingBefore = null, double? spacingAfter = null, PdfDrawingStyle? style = null, string? linkUri = null, string? linkContents = null) {
        Guard.NotNull(drawing, nameof(drawing));
        Guard.Positive(drawing.Width, nameof(drawing.Width));
        Guard.Positive(drawing.Height, nameof(drawing.Height));
        Guard.OptionalUriAction(linkUri, nameof(linkUri));
        if (drawing.Elements.Count == 0) {
            throw new System.ArgumentException("Drawing scenes require at least one shape or text element.", nameof(drawing));
        }

        for (int i = 0; i < drawing.Shapes.Count; i++) {
            var item = drawing.Shapes[i];
            Guard.NotNull(item, nameof(drawing.Shapes));
            Guard.NonNegative(item.X, nameof(drawing.Shapes));
            Guard.NonNegative(item.Y, nameof(drawing.Shapes));
            CreateShapeBlock(item.Shape, PdfAlign.Left, 0, 0);

            if (item.X + item.Shape.Width > drawing.Width || item.Y + item.Shape.Height > drawing.Height) {
                throw new System.ArgumentOutOfRangeException(nameof(drawing), "Drawing scene shapes must fit inside the drawing width and height.");
            }
        }

        for (int i = 0; i < drawing.Elements.Count; i++) {
            var text = drawing.Elements[i] as OfficeDrawingText;
            if (text == null) {
                continue;
            }

            Guard.NotNull(text.Text, nameof(drawing.Elements));
            Guard.NonNegative(text.X, nameof(drawing.Elements));
            Guard.NonNegative(text.Y, nameof(drawing.Elements));
            Guard.Positive(text.Width, nameof(drawing.Elements));
            Guard.Positive(text.Height, nameof(drawing.Elements));
            Guard.Positive(text.Font.Size, nameof(text.Font.Size));
            if (text.LineHeight.HasValue) {
                Guard.Positive(text.LineHeight.Value, nameof(text.LineHeight));
            }

            if (text.X + text.Width > drawing.Width || text.Y + text.Height > drawing.Height) {
                throw new System.ArgumentOutOfRangeException(nameof(drawing), "Drawing scene text must fit inside the drawing width and height.");
            }
        }

        PdfDrawingStyle? drawingStyle = CreateDrawingStyle(align, spacingBefore, spacingAfter, style, "Drawing");
        if (drawingStyle != null) {
            ValidateDrawingStyle(drawingStyle, "Drawing");
        }

        return new DrawingBlock(drawing, drawingStyle, linkUri, linkContents);
    }
}
