using System.Globalization;
using Avalonia;
using Avalonia.Media;
using Avalonia.Media.Imaging;
using OfficeIMO.Drawing;

namespace OfficeIMO.Studio.Features.Reader;

/// <summary>Maps the dependency-free OfficeIMO drawing scene onto Avalonia drawing primitives.</summary>
internal sealed class OfficeDrawingAvaloniaRenderer : IDisposable {
    private readonly Dictionary<OfficeDrawingImage, Bitmap> _images = new(ReferenceEqualityComparer.Instance);

    internal static bool RequiresRasterFallback(OfficeDrawing drawing) => AnalyzeRasterFallback(drawing).Count > 0;

    internal static IReadOnlyList<string> AnalyzeRasterFallback(OfficeDrawing drawing) {
        var reasons = new HashSet<string>(StringComparer.Ordinal);
        AnalyzeRasterFallback(drawing, reasons);
        return reasons.ToArray();
    }

    private static void AnalyzeRasterFallback(OfficeDrawing drawing, HashSet<string> reasons) {
        if (drawing.Fonts.Faces.Count > 0) {
            reasons.Add("Avalonia vector fallback: drawing-scoped embedded fonts require the OfficeIMO raster renderer for glyph fidelity.");
        }
        foreach (OfficeDrawingElement element in drawing.Elements) {
            switch (element) {
                case OfficeDrawingShape shape:
                    AnalyzeRasterFallback(shape.Shape, reasons);
                    break;
                case OfficeDrawingText text:
                    if (text.StackedText || text.ShrinkToFit || text.TextAdvanceWidth.HasValue ||
                        text.HasPadding || text.HasParagraphIndent ||
                        text.UnderlineStyle != OfficeTextDecorationStyle.None ||
                        text.StrikethroughStyle != OfficeTextDecorationStyle.None) {
                        reasons.Add("Avalonia vector fallback: advanced PDF text metrics or decoration require the OfficeIMO raster renderer.");
                    }
                    break;
                case OfficeDrawingImage:
                    break;
                case OfficeDrawingGroup group:
                    AnalyzeRasterFallback(group.Drawing, reasons);
                    break;
                case OfficeDrawingEffectGroup effectGroup:
                    if (effectGroup.BlendMode != OfficeBlendMode.Normal || effectGroup.SoftMask is not null) {
                        reasons.Add("Avalonia vector fallback: blend modes or soft masks require the OfficeIMO raster renderer.");
                    }
                    AnalyzeRasterFallback(effectGroup.Drawing, reasons);
                    break;
                default:
                    reasons.Add($"Avalonia vector fallback: {element.GetType().Name} is not supported by the retained adapter.");
                    break;
            }
        }
    }

    internal void Render(DrawingContext context, OfficeDrawing drawing) {
        ArgumentNullException.ThrowIfNull(context);
        ArgumentNullException.ThrowIfNull(drawing);
        RenderDrawing(context, drawing);
    }

    public void Dispose() {
        ClearImages();
    }

    internal void ClearImages() {
        foreach (Bitmap image in _images.Values) image.Dispose();
        _images.Clear();
    }

    private void RenderDrawing(DrawingContext context, OfficeDrawing drawing) {
        foreach (OfficeDrawingElement element in drawing.Elements) {
            switch (element) {
                case OfficeDrawingShape shape:
                    RenderShape(context, shape);
                    break;
                case OfficeDrawingText text:
                    RenderText(context, text);
                    break;
                case OfficeDrawingImage image:
                    RenderImage(context, image);
                    break;
                case OfficeDrawingGroup group:
                    RenderGroup(context, group);
                    break;
                case OfficeDrawingEffectGroup effectGroup:
                    RenderEffectGroup(context, effectGroup);
                    break;
            }
        }
    }

    private static void AnalyzeRasterFallback(OfficeShape shape, HashSet<string> reasons) {
        if (shape.FillGradient is not null || shape.FillRadialGradient is not null ||
            shape.StrokeGradient is not null || shape.StrokeRadialGradient is not null) {
            reasons.Add("Avalonia vector fallback: gradients require the OfficeIMO raster renderer.");
        }
        if (shape.Shadow is not null || shape.Glow is not null) {
            reasons.Add("Avalonia vector fallback: shadows or glow effects require the OfficeIMO raster renderer.");
        }
        if (shape.StrokeStartMarker is not null || shape.StrokeEndMarker is not null) {
            reasons.Add("Avalonia vector fallback: path markers require the OfficeIMO raster renderer.");
        }
    }

    private static void RenderShape(DrawingContext context, OfficeDrawingShape positioned) {
        OfficeShape shape = positioned.Shape;
        using IDisposable translation = context.PushTransform(Matrix.CreateTranslation(positioned.X, positioned.Y));
        IDisposable? transform = shape.Transform.HasValue
            ? context.PushTransform(ToMatrix(shape.Transform.Value))
            : null;
        IDisposable? clip = shape.ClipPath is not null
            ? context.PushGeometryClip(CreateClipGeometry(shape.ClipPath))
            : null;
        try {
            IBrush? fill = CreateBrush(shape.FillColor, shape.FillOpacity);
            Pen? pen = CreatePen(shape);
            switch (shape.Kind) {
                case OfficeShapeKind.Rectangle:
                    context.DrawRectangle(fill, pen, new Rect(0, 0, shape.Width, shape.Height));
                    break;
                case OfficeShapeKind.RoundedRectangle:
                    context.DrawRectangle(fill, pen, new Rect(0, 0, shape.Width, shape.Height), shape.CornerRadius, shape.CornerRadius);
                    break;
                case OfficeShapeKind.Ellipse:
                    context.DrawEllipse(fill, pen, new Rect(0, 0, shape.Width, shape.Height));
                    break;
                case OfficeShapeKind.Line when shape.Points.Count >= 2:
                    if (pen is not null) context.DrawLine(pen, ToPoint(shape.Points[0]), ToPoint(shape.Points[1]));
                    break;
                case OfficeShapeKind.Polygon:
                    context.DrawGeometry(fill, pen, CreatePolygonGeometry(shape.Points));
                    break;
                case OfficeShapeKind.Path:
                    context.DrawGeometry(fill, pen, CreatePathGeometry(shape.PathCommands, shape.FillRule));
                    break;
            }
        } finally {
            clip?.Dispose();
            transform?.Dispose();
        }
    }

    private static void RenderText(DrawingContext context, OfficeDrawingText text) {
        string family = string.IsNullOrWhiteSpace(text.Font.FamilyName) ? "Arial" : text.Font.FamilyName;
        var typeface = new Typeface(
            family,
            text.Font.IsItalic ? FontStyle.Italic : FontStyle.Normal,
            text.Font.IsBold ? FontWeight.Bold : FontWeight.Normal,
            FontStretch.Normal);
        var formatted = new FormattedText(
            text.Text,
            CultureInfo.CurrentUICulture,
            FlowDirection.LeftToRight,
            typeface,
            Math.Max(1D, text.Font.Size * text.BaselineScale),
            CreateBrush(text.Color ?? OfficeColor.Black, 1D)!) {
            MaxTextWidth = Math.Max(1D, text.Width),
            MaxTextHeight = Math.Max(1D, text.Height),
            TextAlignment = text.Alignment switch {
                OfficeTextAlignment.Center => TextAlignment.Center,
                OfficeTextAlignment.Right => TextAlignment.Right,
                OfficeTextAlignment.Justify => TextAlignment.Justify,
                _ => TextAlignment.Left
            }
        };
        if (text.LineHeight.HasValue) formatted.LineHeight = text.LineHeight.Value;

        double y = text.Y + text.BaselineOffset;
        if (text.VerticalAlignment == OfficeTextVerticalAlignment.Center) y += Math.Max(0D, (text.Height - formatted.Height) / 2D);
        if (text.VerticalAlignment == OfficeTextVerticalAlignment.Bottom) y += Math.Max(0D, text.Height - formatted.Height);

        IDisposable? transform = text.HasFrameTransform
            ? context.PushTransform(ToMatrix(text.CreateFrameTransform().CreateDestinationTransform()))
            : null;
        try {
            context.DrawText(formatted, new Point(text.X, y));
        } finally {
            transform?.Dispose();
        }
    }

    private void RenderImage(DrawingContext context, OfficeDrawingImage image) {
        if (!_images.TryGetValue(image, out Bitmap? bitmap)) {
            using var stream = new MemoryStream(image.Bytes, writable: false);
            bitmap = new Bitmap(stream);
            _images.Add(image, bitmap);
        }

        OfficeImageProjection projection = image.Projection;
        Rect source = projection.HasCrop
            ? new Rect(
                bitmap.PixelSize.Width * projection.SourceLeft,
                bitmap.PixelSize.Height * projection.SourceTop,
                bitmap.PixelSize.Width * projection.SourceWidth,
                bitmap.PixelSize.Height * projection.SourceHeight)
            : new Rect(0, 0, bitmap.PixelSize.Width, bitmap.PixelSize.Height);
        Rect destination = new(projection.X, projection.Y, projection.Width, projection.Height);
        IDisposable? opacity = image.Opacity < 0.999999D ? context.PushOpacity(image.Opacity) : null;
        IDisposable? transform = projection.HasTransform
            ? context.PushTransform(ToMatrix(projection.CreateFrameTransform().CreateDestinationTransform()))
            : null;
        try {
            context.DrawImage(bitmap, source, destination);
        } finally {
            transform?.Dispose();
            opacity?.Dispose();
        }
    }

    private void RenderGroup(DrawingContext context, OfficeDrawingGroup group) {
        using IDisposable translation = context.PushTransform(Matrix.CreateTranslation(group.X, group.Y));
        IDisposable? transform = group.FrameTransform.HasValue
            ? context.PushTransform(ToMatrix(group.FrameTransform.Value.CreateDestinationTransform()))
            : null;
        using IDisposable clip = context.PushGeometryClip(CreateClipGeometry(group.ClipPath));
        using IDisposable content = context.PushTransform(Matrix.CreateTranslation(group.ContentOffsetX, group.ContentOffsetY));
        try {
            RenderDrawing(context, group.Drawing);
        } finally {
            transform?.Dispose();
        }
    }

    private void RenderEffectGroup(DrawingContext context, OfficeDrawingEffectGroup group) {
        using IDisposable transform = context.PushTransform(ToMatrix(group.Transform));
        IDisposable? opacity = group.Opacity < 0.999999D ? context.PushOpacity(group.Opacity) : null;
        try {
            RenderDrawing(context, group.Drawing);
        } finally {
            opacity?.Dispose();
        }
    }

    private static IBrush? CreateBrush(OfficeColor? color, double? opacity) {
        if (!color.HasValue) return null;
        OfficeColor value = color.Value;
        double combined = Math.Clamp(opacity ?? 1D, 0D, 1D) * value.A / 255D;
        return new SolidColorBrush(Color.FromArgb(
            (byte)Math.Round(combined * 255D),
            value.R,
            value.G,
            value.B));
    }

    private static Pen? CreatePen(OfficeShape shape) {
        IBrush? brush = CreateBrush(shape.StrokeColor, shape.StrokeOpacity);
        if (brush is null || shape.StrokeWidth <= 0D) return null;
        IDashStyle? dashStyle = shape.StrokeDashStyle switch {
            OfficeStrokeDashStyle.Dash => DashStyle.Dash,
            OfficeStrokeDashStyle.Dot => DashStyle.Dot,
            OfficeStrokeDashStyle.DashDot => DashStyle.DashDot,
            OfficeStrokeDashStyle.DashDotDot => DashStyle.DashDotDot,
            _ => null
        };
        return new Pen(
            brush,
            shape.StrokeWidth,
            dashStyle,
            shape.StrokeLineCap switch {
                OfficeStrokeLineCap.Round => PenLineCap.Round,
                OfficeStrokeLineCap.Square => PenLineCap.Square,
                _ => PenLineCap.Flat
            },
            shape.StrokeLineJoin switch {
                OfficeStrokeLineJoin.Round => PenLineJoin.Round,
                OfficeStrokeLineJoin.Bevel => PenLineJoin.Bevel,
                _ => PenLineJoin.Miter
            });
    }

    private static Geometry CreatePolygonGeometry(IReadOnlyList<OfficePoint> points) {
        if (points.Count == 0) return StreamGeometry.Parse("M 0,0");
        var geometry = new StreamGeometry();
        using StreamGeometryContext context = geometry.Open();
        context.BeginFigure(ToPoint(points[0]), true);
        for (int i = 1; i < points.Count; i++) context.LineTo(ToPoint(points[i]));
        context.EndFigure(true);
        return geometry;
    }

    private static Geometry CreatePathGeometry(IReadOnlyList<OfficePathCommand> commands, OfficeFillRule fillRule) {
        var geometry = new StreamGeometry();
        using StreamGeometryContext context = geometry.Open();
        context.SetFillRule(fillRule == OfficeFillRule.NonZero ? FillRule.NonZero : FillRule.EvenOdd);
        bool figureOpen = false;
        foreach (OfficePathCommand command in commands) {
            switch (command.Kind) {
                case OfficePathCommandKind.MoveTo:
                    if (figureOpen) context.EndFigure(false);
                    context.BeginFigure(ToPoint(command.Point), true);
                    figureOpen = true;
                    break;
                case OfficePathCommandKind.LineTo when figureOpen:
                    context.LineTo(ToPoint(command.Point));
                    break;
                case OfficePathCommandKind.QuadraticBezierTo when figureOpen:
                    context.QuadraticBezierTo(ToPoint(command.ControlPoint1), ToPoint(command.Point));
                    break;
                case OfficePathCommandKind.CubicBezierTo when figureOpen:
                    context.CubicBezierTo(ToPoint(command.ControlPoint1), ToPoint(command.ControlPoint2), ToPoint(command.Point));
                    break;
                case OfficePathCommandKind.Close when figureOpen:
                    context.EndFigure(true);
                    figureOpen = false;
                    break;
            }
        }
        if (figureOpen) context.EndFigure(false);
        return geometry;
    }

    private static Geometry CreateClipGeometry(OfficeClipPath clip) => clip.Kind switch {
        OfficeClipPathKind.Rectangle => new RectangleGeometry(new Rect(0, 0, clip.Width, clip.Height)),
        OfficeClipPathKind.RoundedRectangle => new RectangleGeometry(new Rect(0, 0, clip.Width, clip.Height), clip.CornerRadius, clip.CornerRadius),
        OfficeClipPathKind.Path => CreatePathGeometry(clip.Commands, clip.FillRule),
        _ => new RectangleGeometry(new Rect(0, 0, 0, 0))
    };

    private static Matrix ToMatrix(OfficeTransform transform) => new(
        transform.M11,
        transform.M12,
        transform.M21,
        transform.M22,
        transform.OffsetX,
        transform.OffsetY);

    private static Point ToPoint(OfficePoint point) => new(point.X, point.Y);
}
