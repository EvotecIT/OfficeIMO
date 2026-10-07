using System;
using System.Collections.Generic;
#if NET8_0_OR_GREATER
using System.Buffers;
#endif

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free drawing canvas for an <see cref="OfficeRasterImage"/>.
/// </summary>
public sealed partial class OfficeRasterCanvas {
    private const int AntiAliasSamples = 3;
    private const int MaximumContourCoverageTileWidth = 8192;
    private const double MinimumDashSegmentAdvance = 1E-9D;
    private const double MinimumRasterDashLength = 0.25D;
    private static readonly Lazy<OfficeTrueTypeFont?> DefaultFont = new Lazy<OfficeTrueTypeFont?>(OfficeTrueTypeFont.TryLoadDefault);
    private readonly OfficeRasterImage? _image;
    private readonly OfficeRasterRenderTarget? _target;
    private readonly OfficeTrueTypeFont? _font;
    private readonly OfficeFontFaceCollection? _fonts;
    private readonly bool _scopedFontResolutionOnly;
    private readonly IOfficeTextShapingProvider? _textShapingProvider;
    private readonly string? _textShapingLanguage;
    private readonly ICollection<OfficeImageExportDiagnostic>? _diagnosticSink;
    private readonly string? _diagnosticSource;
    private readonly System.Threading.CancellationToken _cancellationToken;
    private bool _reportedBoundedTextShapingFallback;
    private bool _reportedIncompleteTextShapingFallback;
    private const long MaximumTransformedTextIntermediatePixels = 64_000_000L;
    private OfficeRasterTransformedTextBudget _transformedTextBudget = new OfficeRasterTransformedTextBudget();
    private int CoverageSamples => _target != null && _target.Supersampling > 1 ? 1 : AntiAliasSamples;

    private static bool IsFinite(double value) => !double.IsNaN(value) && !double.IsInfinity(value);

    /// <summary>
    /// Creates a canvas over the supplied image.
    /// </summary>
    public OfficeRasterCanvas(
        OfficeRasterImage image,
        OfficeTrueTypeFont? font = null,
        OfficeFontFaceCollection? fonts = null)
        : this(
            image,
            font,
            fonts,
            textShapingProvider: null,
            textShapingLanguage: null,
            diagnosticSink: null,
            diagnosticSource: null,
            cancellationToken: default) {
    }

    /// <summary>Creates a canvas with complex-text shaping and fidelity diagnostics.</summary>
    public OfficeRasterCanvas(
        OfficeRasterImage image,
        OfficeTrueTypeFont? font,
        OfficeFontFaceCollection? fonts,
        IOfficeTextShapingProvider? textShapingProvider = null,
        string? textShapingLanguage = null,
        ICollection<OfficeImageExportDiagnostic>? diagnosticSink = null,
        string? diagnosticSource = null,
        System.Threading.CancellationToken cancellationToken = default) {
        _image = image ?? throw new ArgumentNullException(nameof(image));
        _font = font ?? DefaultFont.Value;
        _fonts = fonts?.Clone();
        _textShapingProvider = textShapingProvider;
        _textShapingLanguage = NormalizeTextShapingLanguage(textShapingLanguage);
        _diagnosticSink = diagnosticSink;
        _diagnosticSource = diagnosticSource;
        _cancellationToken = cancellationToken;
    }

    /// <summary>
    /// Creates a canvas over the supplied supersampled render target.
    /// </summary>
    public OfficeRasterCanvas(
        OfficeRasterRenderTarget target,
        OfficeTrueTypeFont? font = null,
        OfficeFontFaceCollection? fonts = null)
        : this(
            target,
            font,
            fonts,
            textShapingProvider: null,
            textShapingLanguage: null,
            diagnosticSink: null,
            diagnosticSource: null,
            cancellationToken: default) {
    }

    /// <summary>Creates a supersampled canvas with complex-text shaping and fidelity diagnostics.</summary>
    public OfficeRasterCanvas(
        OfficeRasterRenderTarget target,
        OfficeTrueTypeFont? font,
        OfficeFontFaceCollection? fonts,
        IOfficeTextShapingProvider? textShapingProvider = null,
        string? textShapingLanguage = null,
        ICollection<OfficeImageExportDiagnostic>? diagnosticSink = null,
        string? diagnosticSource = null,
        System.Threading.CancellationToken cancellationToken = default) {
        _target = target ?? throw new ArgumentNullException(nameof(target));
        _font = font ?? DefaultFont.Value;
        _fonts = fonts?.Clone();
        _textShapingProvider = textShapingProvider;
        _textShapingLanguage = NormalizeTextShapingLanguage(textShapingLanguage);
        _diagnosticSink = diagnosticSink;
        _diagnosticSource = diagnosticSource;
        _cancellationToken = cancellationToken;
    }

    private static string? NormalizeTextShapingLanguage(string? value) =>
        string.IsNullOrWhiteSpace(value) ? null : value!.Trim();

    /// <summary>Canvas width in pixels.</summary>
    public int Width => _image?.Width ?? _target!.RenderWidth;

    /// <summary>Canvas height in pixels.</summary>
    public int Height => _image?.Height ?? _target!.RenderHeight;

    internal IOfficeTextShapingProvider? TextShapingProvider => _textShapingProvider;

    internal string? TextShapingLanguage => _textShapingLanguage;

    internal OfficeTrueTypeFont? OutlineFont => _font;

    internal OfficeFontFaceCollection? Fonts => _fonts;

    internal void ChargeTransformedTextIntermediatePixels(long pixels, long maximumRasterPixels) {
        long consumed = _transformedTextBudget.Pixels;
        if (pixels < 0L || pixels > MaximumTransformedTextIntermediatePixels - consumed) {
            throw new OfficeImageExportLimitException(1D,
                pixels > long.MaxValue - consumed ? long.MaxValue : consumed + pixels,
                MaximumTransformedTextIntermediatePixels,
                OfficeRasterImageEncoder.GetMaximumDimension(OfficeImageExportFormat.Png));
        }
        _transformedTextBudget.ChargeIntermediateSurfacePixels(pixels, maximumRasterPixels);
        _transformedTextBudget.Pixels = consumed + pixels;
    }

    internal void ReleaseTransformedTextIntermediatePixels(long pixels) {
        _transformedTextBudget.Pixels -= pixels;
        _transformedTextBudget.ReleaseIntermediateSurfacePixels(pixels);
    }

    internal OfficeRasterTransformedTextBudget TransformedTextBudget => _transformedTextBudget;

    internal void ChargeIntermediateSurfacePixels(long pixels, long maximumRasterPixels) =>
        _transformedTextBudget.ChargeIntermediateSurfacePixels(pixels, maximumRasterPixels);

    internal void ShareTransformedTextBudget(OfficeRasterTransformedTextBudget budget) =>
        _transformedTextBudget = budget;

    internal System.Threading.CancellationToken CancellationToken => _cancellationToken;

    internal ICollection<OfficeImageExportDiagnostic>? DiagnosticSink => _diagnosticSink;

    internal string? DiagnosticSource => _diagnosticSource;

    /// <summary>Fills a rectangle.</summary>
    public void FillRectangle(double x, double y, double width, double height, OfficeColor color) {
        if (width <= 0D || height <= 0D || color.A == 0) return;
        FillPolygonCore(RectanglePoints(x, y, width, height), color);
    }

    /// <summary>Fills a rectangle with a linear gradient.</summary>
    public void FillLinearGradientRectangle(double x, double y, double width, double height, OfficeLinearGradient gradient) {
        if (gradient == null) throw new ArgumentNullException(nameof(gradient));
        if (width <= 0D || height <= 0D) return;
        FillPolygonCore(RectanglePoints(x, y, width, height), gradient);
    }

    /// <summary>Fills a rectangle with a radial gradient.</summary>
    public void FillRadialGradientRectangle(double x, double y, double width, double height, OfficeRadialGradient gradient) {
        if (gradient == null) throw new ArgumentNullException(nameof(gradient));
        if (width <= 0D || height <= 0D) return;
        FillPolygonCore(RectanglePoints(x, y, width, height), gradient);
    }

    /// <summary>Draws a rectangle outline.</summary>
    public void DrawRectangle(double x, double y, double width, double height, OfficeColor color, double thickness = 1D) {
        if (width <= 0D || height <= 0D) return;
        StrokePolyline(new[] { new OfficePoint(x, y), new OfficePoint(x + width, y), new OfficePoint(x + width, y + height), new OfficePoint(x, y + height) }, color, thickness, closed: true);
    }

    /// <summary>Fills an ellipse bounded by the supplied rectangle.</summary>
    public void FillEllipse(double x, double y, double width, double height, OfficeColor color) {
        if (width <= 0D || height <= 0D || color.A == 0) return;
        FillPolygonCore(OfficeCurveFlattening.Ellipse(x + width / 2D, y + height / 2D, width / 2D, height / 2D, 1D), color);
    }

    /// <summary>Draws an ellipse outline bounded by the supplied rectangle.</summary>
    public void DrawEllipse(double x, double y, double width, double height, OfficeColor color, double thickness = 1D) {
        if (width <= 0D || height <= 0D) return;
        StrokePolyline(OfficeCurveFlattening.Ellipse(x + width / 2D, y + height / 2D, width / 2D, height / 2D, 1D), color, thickness, closed: true);
    }

    /// <summary>Draws a filled and/or stroked ellipse using center/radius coordinates and optional rotation.</summary>
    public void DrawEllipse(
        double centerX,
        double centerY,
        double radiusX,
        double radiusY,
        OfficeColor fill,
        OfficeColor stroke,
        double thickness = 1D,
        double rotationDegrees = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D) {
        if (radiusX <= 0D || radiusY <= 0D) return;
        var points = OfficeCurveFlattening.Ellipse(centerX, centerY, radiusX, radiusY, 1D);
        double radians = OfficeGeometry.DegreesToRadians(rotationDegrees);
        for (int i = 0; i < points.Count; i++) points[i] = OfficeGeometry.RotatePoint(points[i], rotationCenterX, rotationCenterY, radians);
        if (fill.A > 0) FillPolygonCore(points, fill);
        StrokePolyline(points, stroke, thickness, closed: true);
    }

    /// <summary>Draws a line segment.</summary>
    public void DrawLine(double x1, double y1, double x2, double y2, OfficeColor color, double thickness = 1D) {
        if (color.A == 0 || thickness <= 0D) {
            return;
        }

        DrawLineSegment(x1, y1, x2, y2, color, thickness);
    }

    /// <summary>Draws a dashed line segment.</summary>
    public void DrawDashedLine(double x1, double y1, double x2, double y2, OfficeColor color, double thickness = 1D, double dashLength = 6D, double gapLength = 4D) {
        if (!IsFinite(dashLength) || !IsFinite(gapLength) || dashLength <= 0D || gapLength < 0D) return;
        NormalizeRasterDashLengths(ref dashLength, ref gapLength);
        StrokePolyline(new[] { new OfficePoint(x1, y1), new OfficePoint(x2, y2) }, color, thickness, pattern: new[] { dashLength, gapLength });
    }

    /// <summary>Draws a line segment using a shared Office stroke dash style.</summary>
    public void DrawStyledLine(double x1, double y1, double x2, double y2, OfficeColor color, double thickness = 1D, OfficeStrokeDashStyle dashStyle = OfficeStrokeDashStyle.Solid) {
        if (dashStyle == OfficeStrokeDashStyle.Solid) {
            DrawLine(x1, y1, x2, y2, color, thickness);
            return;
        }

        DrawPatternedLine(x1, y1, x2, y2, color, thickness, dashStyle.GetDashPattern(thickness));
    }

    /// <summary>
    /// Draws two parallel line segments using a shared Office stroke dash style.
    /// </summary>
    /// <param name="x1">Source line start X coordinate.</param>
    /// <param name="y1">Source line start Y coordinate.</param>
    /// <param name="x2">Source line end X coordinate.</param>
    /// <param name="y2">Source line end Y coordinate.</param>
    /// <param name="color">Stroke color.</param>
    /// <param name="thickness">Stroke thickness.</param>
    /// <param name="separation">Distance between the two parallel line centers.</param>
    /// <param name="dashStyle">Stroke dash style.</param>
    public void DrawParallelStyledLine(double x1, double y1, double x2, double y2, OfficeColor color, double thickness, double separation, OfficeStrokeDashStyle dashStyle = OfficeStrokeDashStyle.Solid) {
        if (!OfficeGeometry.TryGetParallelLineOffsets(x1, y1, x2, y2, separation, out double offsetX, out double offsetY)) {
            return;
        }

        DrawStyledLine(x1 - offsetX, y1 - offsetY, x2 - offsetX, y2 - offsetY, color, thickness, dashStyle);
        DrawStyledLine(x1 + offsetX, y1 + offsetY, x2 + offsetX, y2 + offsetY, color, thickness, dashStyle);
    }

    /// <summary>Draws a line segment using an alternating dash and gap pattern.</summary>
    public void DrawPatternedLine(double x1, double y1, double x2, double y2, OfficeColor color, double thickness, IReadOnlyList<double>? dashPattern) {
        StrokePolyline(new[] { new OfficePoint(x1, y1), new OfficePoint(x2, y2) }, color, thickness, pattern: dashPattern);
    }

    /// <summary>Draws an elliptical arc using center/radius coordinates and optional rotation.</summary>
    public void DrawArc(
        double centerX,
        double centerY,
        double radiusX,
        double radiusY,
        double startDegrees,
        double endDegrees,
        OfficeColor color,
        double thickness = 1D,
        double rotationDegrees = 0D,
        double rotationCenterX = 0D,
        double rotationCenterY = 0D) {
        if (color.A == 0 || thickness <= 0D || radiusX <= 0D || radiusY <= 0D) return;
        double start = OfficeGeometry.DegreesToRadians(startDegrees), sweep = OfficeGeometry.DegreesToRadians(endDegrees - startDegrees);
        if (!IsFinite(sweep) || Math.Abs(sweep) <= 1E-9D) return;
        int segments = OfficeCurveFlattening.ArcSegments(Math.Max(radiusX, radiusY), sweep);
        double rotation = OfficeGeometry.DegreesToRadians(rotationDegrees);
        var points = new List<OfficePoint> { CreateArcStartPoint(centerX, centerY, radiusX, radiusY, start, rotation, rotationCenterX, rotationCenterY) };
        points.AddRange(OfficeGeometry.CreateEllipticalArcPoints(centerX, centerY, radiusX, radiusY, start, sweep, segments, rotation, rotationCenterX, rotationCenterY));
        StrokePolyline(points, color, thickness);
    }

    private void DrawLineSegment(double x1, double y1, double x2, double y2, OfficeColor color, double thickness) =>
        StrokePolyline(new[] { new OfficePoint(x1, y1), new OfficePoint(x2, y2) }, color, thickness);
    /// <summary>Fills a polygon.</summary>
    public void FillPolygon(IReadOnlyList<OfficePoint> points, OfficeColor color) {
        if (color.A == 0 || points == null || points.Count < 3) {
            return;
        }

        FillPolygonCore(points, color);
    }

    /// <summary>Fills a polygon with a linear gradient fitted to the polygon bounds.</summary>
    public void FillLinearGradientPolygon(IReadOnlyList<OfficePoint> points, OfficeLinearGradient gradient) {
        if (gradient == null) {
            throw new ArgumentNullException(nameof(gradient));
        }

        if (points == null || points.Count < 3) {
            return;
        }

        FillPolygonCore(points, gradient);
    }

    /// <summary>Fills a polygon with a radial gradient.</summary>
    public void FillRadialGradientPolygon(IReadOnlyList<OfficePoint> points, OfficeRadialGradient gradient) {
        if (points == null) {
            throw new ArgumentNullException(nameof(points));
        }

        if (gradient == null) {
            throw new ArgumentNullException(nameof(gradient));
        }

        if (points.Count < 3) {
            return;
        }

        FillPolygonCore(points, gradient);
    }

    /// <summary>Fills multiple polygon contours using the even-odd fill rule.</summary>
    public void FillPolygonsEvenOdd(IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeColor color) {
        if (color.A == 0 || contours == null || contours.Count == 0) {
            return;
        }

        FillContours(contours, color, OfficeFillRule.EvenOdd);
    }

    /// <summary>Fills multiple polygon contours using the non-zero winding fill rule.</summary>
    public void FillPolygonsNonZero(IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeColor color) {
        if (color.A == 0 || contours == null || contours.Count == 0) {
            return;
        }

        FillContours(contours, color, OfficeFillRule.NonZero);
    }

    /// <summary>Strokes a polygon outline.</summary>
    public void DrawPolygon(IReadOnlyList<OfficePoint> points, OfficeColor color, double thickness = 1D) {
        if (color.A == 0 || points == null || points.Count < 2) {
            return;
        }

        StrokePolyline(points, color, thickness, closed: points.Count > 2);
    }
    /// <summary>Draws an image scaled into the supplied rectangle.</summary>
    public void DrawImage(OfficeRasterImage image, double x, double y, double width, double height) {
        DrawImage(
            image,
            x,
            y,
            width,
            height,
            sourceLeft: 0D,
            sourceTop: 0D,
            sourceWidth: 1D,
            sourceHeight: 1D,
            rotationDegrees: 0D,
            rotationCenterX: x + (width / 2D),
            rotationCenterY: y + (height / 2D),
            flipHorizontal: false,
            flipVertical: false);
    }

    /// <summary>
    /// Draws an image using a shared projection that carries placement, source crop, rotation, and flips.
    /// </summary>
    /// <param name="image">Image to draw.</param>
    /// <param name="projection">Shared image projection.</param>
    public void DrawImage(OfficeRasterImage image, OfficeImageProjection projection) {
        DrawImage(image, projection, interpolate: true);
    }

    /// <summary>Draws an image using a shared projection and the requested sampling behavior.</summary>
    public void DrawImage(OfficeRasterImage image, OfficeImageProjection projection, bool interpolate) {
        DrawImage(
            image,
            projection.X,
            projection.Y,
            projection.Width,
            projection.Height,
            projection.SourceLeft,
            projection.SourceTop,
            projection.SourceWidth,
            projection.SourceHeight,
            projection.RotationDegrees,
            projection.RotationCenterX,
            projection.RotationCenterY,
            projection.FlipHorizontal,
            projection.FlipVertical,
            interpolate);
    }

    /// <summary>
    /// Draws a source rectangle from an image scaled into the supplied destination rectangle.
    /// Source coordinates are normalized, where 0 is the left/top edge and 1 is the right/bottom edge.
    /// </summary>
    public void DrawImage(OfficeRasterImage image, double x, double y, double width, double height, double sourceLeft, double sourceTop, double sourceWidth, double sourceHeight) {
        DrawImage(
            image,
            x,
            y,
            width,
            height,
            sourceLeft,
            sourceTop,
            sourceWidth,
            sourceHeight,
            rotationDegrees: 0D,
            rotationCenterX: x + (width / 2D),
            rotationCenterY: y + (height / 2D),
            flipHorizontal: false,
            flipVertical: false);
    }

    /// <summary>Draws an image scaled and rotated around the supplied rotation center.</summary>
    public void DrawImage(OfficeRasterImage image, double x, double y, double width, double height, double rotationDegrees, double rotationCenterX, double rotationCenterY) {
        DrawImage(
            image,
            x,
            y,
            width,
            height,
            sourceLeft: 0D,
            sourceTop: 0D,
            sourceWidth: 1D,
            sourceHeight: 1D,
            rotationDegrees,
            rotationCenterX,
            rotationCenterY,
            flipHorizontal: false,
            flipVertical: false);
    }

    /// <summary>
    /// Draws a source rectangle from an image into the supplied destination rectangle with optional rotation and flips.
    /// Source coordinates are normalized, where 0 is the left/top edge and 1 is the right/bottom edge.
    /// </summary>
    public void DrawImage(
        OfficeRasterImage image,
        double x,
        double y,
        double width,
        double height,
        double sourceLeft,
        double sourceTop,
        double sourceWidth,
        double sourceHeight,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY,
        bool flipHorizontal,
        bool flipVertical) {
        DrawImage(image, x, y, width, height, sourceLeft, sourceTop, sourceWidth, sourceHeight,
            rotationDegrees, rotationCenterX, rotationCenterY, flipHorizontal, flipVertical, interpolate: true);
    }

    /// <summary>Draws a transformed image with explicit scaling interpolation behavior.</summary>
    public void DrawImage(
        OfficeRasterImage image,
        double x,
        double y,
        double width,
        double height,
        double sourceLeft,
        double sourceTop,
        double sourceWidth,
        double sourceHeight,
        double rotationDegrees,
        double rotationCenterX,
        double rotationCenterY,
        bool flipHorizontal,
        bool flipVertical,
        bool interpolate) {
        if (image == null || width <= 0D || height <= 0D) {
            return;
        }

        sourceLeft = Clamp(sourceLeft, 0D, 1D);
        sourceTop = Clamp(sourceTop, 0D, 1D);
        sourceWidth = Math.Min(Math.Max(0D, sourceWidth), 1D - sourceLeft);
        sourceHeight = Math.Min(Math.Max(0D, sourceHeight), 1D - sourceTop);
        if (sourceWidth <= 0D || sourceHeight <= 0D) {
            return;
        }

        var projection = new OfficeImageProjection(
            new OfficeImagePlacement(x, y, width, height),
            new OfficeImageSourceCrop(
                sourceLeft,
                sourceTop,
                Math.Max(0D, 1D - sourceLeft - sourceWidth),
                Math.Max(0D, 1D - sourceTop - sourceHeight)),
            rotationDegrees,
            rotationCenterX,
            rotationCenterY,
            flipHorizontal,
            flipVertical);
        OfficeTransform imageTransform = ScaleCoordinates(projection.CreateUnitSquareTransform());
        if (!imageTransform.TryInvert(out OfficeTransform inverseTransform)) {
            return;
        }

        if (interpolate) image = PrefilterImage(image,
            SamplingAxisLength(inverseTransform.M11, inverseTransform.M21) * image.Width * sourceWidth,
            SamplingAxisLength(inverseTransform.M12, inverseTransform.M22) * image.Height * sourceHeight);

        (double minX, double minY, double maxX, double maxY) = imageTransform.TransformRectangleBounds(0D, 0D, 1D, 1D);
        int left = Clamp((int)Math.Floor(minX), 0, Width - 1);
        int top = Clamp((int)Math.Floor(minY), 0, Height - 1);
        int right = Clamp((int)Math.Ceiling(maxX), 0, Width - 1);
        int bottom = Clamp((int)Math.Ceiling(maxY), 0, Height - 1);
        _cancellationToken.ThrowIfCancellationRequested();
        for (int py = top; py <= bottom; py++) {
            _cancellationToken.ThrowIfCancellationRequested();
            for (int px = left; px <= right; px++) {
                OfficePoint unit = inverseTransform.TransformPoint(new OfficePoint(px + 0.5D, py + 0.5D));
                double u = unit.X;
                double v = unit.Y;
                if (u < 0D || u >= 1D || v < 0D || v >= 1D) {
                    continue;
                }

                double sourceX = ((sourceLeft + (u * sourceWidth)) * image.Width) - 0.5D;
                double sourceY = ((sourceTop + (v * sourceHeight)) * image.Height) - 0.5D;

                BlendPixel(px, py, interpolate
                    ? SampleBilinear(image, sourceX, sourceY)
                    : image.GetPixel(
                        Clamp((int)Math.Floor(sourceX + 0.5D), 0, image.Width - 1),
                        Clamp((int)Math.Floor(sourceY + 0.5D), 0, image.Height - 1)));
            }
        }
        _cancellationToken.ThrowIfCancellationRequested();
    }

    /// <summary>Draws an image through an arbitrary destination-space affine transform.</summary>
    public void DrawAffineImage(OfficeRasterImage image, OfficeTransform transform, double opacity = 1D) =>
        DrawAffineImage(image, transform, opacity, interpolate: true);

    internal void DrawAffineImage(OfficeRasterImage image, OfficeTransform transform, double opacity, bool interpolate) {
        transform = ScaleCoordinates(transform);
        if (image == null) throw new ArgumentNullException(nameof(image));
        if (double.IsNaN(opacity) || double.IsInfinity(opacity) || opacity < 0D || opacity > 1D) {
            throw new ArgumentOutOfRangeException(nameof(opacity), "Image opacity must be between zero and one.");
        }
        if (opacity <= 0D || !transform.TryInvert(out OfficeTransform inverse)) return;

        (double minX, double minY, double maxX, double maxY) = transform.TransformRectangleBounds(0D, 0D, image.Width, image.Height);
        if (interpolate) image = PrefilterAffineImage(image, ref inverse);
        int left = Clamp((int)Math.Floor(minX), 0, Width - 1);
        int top = Clamp((int)Math.Floor(minY), 0, Height - 1);
        int right = Clamp((int)Math.Ceiling(maxX), 0, Width - 1);
        int bottom = Clamp((int)Math.Ceiling(maxY), 0, Height - 1);
        for (int py = top; py <= bottom; py++) {
            _cancellationToken.ThrowIfCancellationRequested();
            for (int px = left; px <= right; px++) {
                OfficePoint source = inverse.TransformPoint(new OfficePoint(px + 0.5D, py + 0.5D));
                if (source.X < 0D || source.X >= image.Width || source.Y < 0D || source.Y >= image.Height) continue;
                OfficeColor color = interpolate
                    ? SampleBilinear(image, source.X - 0.5D, source.Y - 0.5D)
                    : image.GetPixel(
                        Clamp((int)Math.Floor(source.X), 0, image.Width - 1),
                        Clamp((int)Math.Floor(source.Y), 0, image.Height - 1));
                if (opacity < 1D) color = OfficeColor.FromRgba(color.R, color.G, color.B, (byte)Math.Round(color.A * opacity));
                BlendPixel(px, py, color);
            }
        }
    }

    private static bool ContainsPoint(IReadOnlyList<OfficePoint> points, double x, double y) {
        bool inside = false;
        int j = points.Count - 1;
        for (int i = 0; i < points.Count; i++) {
            double xi = points[i].X;
            double yi = points[i].Y;
            double xj = points[j].X;
            double yj = points[j].Y;
            bool intersect = ((yi > y) != (yj > y)) && x < ((xj - xi) * (y - yi) / ((yj - yi) == 0D ? double.Epsilon : (yj - yi)) + xi);
            if (intersect) {
                inside = !inside;
            }

            j = i;
        }

        return inside;
    }

    private static int GetWindingNumber(IReadOnlyList<OfficePoint> points, double x, double y) {
        int winding = 0;
        for (int i = 0, j = points.Count - 1; i < points.Count; j = i++) {
            OfficePoint start = points[j];
            OfficePoint end = points[i];
            if (start.Y <= y) {
                if (end.Y > y && IsLeft(start, end, x, y) > 0D) {
                    winding++;
                }
            } else if (end.Y <= y && IsLeft(start, end, x, y) < 0D) {
                winding--;
            }
        }

        return winding;
    }

    private static double IsLeft(OfficePoint start, OfficePoint end, double x, double y) =>
        ((end.X - start.X) * (y - start.Y)) - ((x - start.X) * (end.Y - start.Y));

    private void BlendPixel(int x, int y, OfficeColor color) {
        _cancellationToken.ThrowIfCancellationRequested();
        if (!IsPixelInsideClip(x, y)) {
            return;
        }

        if (_image != null) {
            _image.BlendPixel(x, y, color);
            return;
        }

        _target!.BlendPixel(x, y, color);
    }

    private bool IsOutsideCanvas(double x, double y, double width, double height) =>
        x >= Width || y >= Height || x + width <= 0D || y + height <= 0D;

    private static OfficeColor SampleBilinear(OfficeRasterImage image, double sourceX, double sourceY) {
        int x0 = Clamp((int)Math.Floor(sourceX), 0, image.Width - 1);
        int y0 = Clamp((int)Math.Floor(sourceY), 0, image.Height - 1);
        int x1 = Clamp(x0 + 1, 0, image.Width - 1);
        int y1 = Clamp(y0 + 1, 0, image.Height - 1);
        double tx = Clamp(sourceX - x0, 0D, 1D);
        double ty = Clamp(sourceY - y0, 0D, 1D);
        OfficeColor c00 = image.GetPixel(x0, y0);
        OfficeColor c10 = image.GetPixel(x1, y0);
        OfficeColor c01 = image.GetPixel(x0, y1);
        OfficeColor c11 = image.GetPixel(x1, y1);
        double w00 = (1D - tx) * (1D - ty);
        double w10 = tx * (1D - ty);
        double w01 = (1D - tx) * ty;
        double w11 = tx * ty;
        double alpha = (c00.A * w00) + (c10.A * w10) + (c01.A * w01) + (c11.A * w11);
        if (alpha <= 0D) {
            return OfficeColor.Transparent;
        }

        return OfficeColor.FromRgba(
            SamplePremultipliedChannel(c00.R, c00.A, w00, c10.R, c10.A, w10, c01.R, c01.A, w01, c11.R, c11.A, w11, alpha),
            SamplePremultipliedChannel(c00.G, c00.A, w00, c10.G, c10.A, w10, c01.G, c01.A, w01, c11.G, c11.A, w11, alpha),
            SamplePremultipliedChannel(c00.B, c00.A, w00, c10.B, c10.A, w10, c01.B, c01.A, w01, c11.B, c11.A, w11, alpha),
            (byte)Math.Round(Clamp(alpha, 0D, 255D)));
    }

    private static byte SamplePremultipliedChannel(
        byte c00,
        byte a00,
        double w00,
        byte c10,
        byte a10,
        double w10,
        byte c01,
        byte a01,
        double w01,
        byte c11,
        byte a11,
        double w11,
        double alpha) {
        double premultiplied =
            (c00 * a00 * w00) +
            (c10 * a10 * w10) +
            (c01 * a01 * w01) +
            (c11 * a11 * w11);
        return (byte)Math.Round(Clamp(premultiplied / alpha, 0D, 255D));
    }

    private static double Distance(double x1, double y1, double x2, double y2) {
        double dx = x2 - x1;
        double dy = y2 - y1;
        return Math.Sqrt((dx * dx) + (dy * dy));
    }

    private static OfficePoint CreateArcStartPoint(double centerX, double centerY, double radiusX, double radiusY, double startRadians, double rotationRadians, double rotationCenterX, double rotationCenterY) {
        OfficePoint point = new OfficePoint(centerX + (Math.Cos(startRadians) * radiusX), centerY + (Math.Sin(startRadians) * radiusY));
        return Math.Abs(rotationRadians) > 0.000001D
            ? OfficeGeometry.RotatePoint(point, rotationCenterX, rotationCenterY, rotationRadians)
            : point;
    }

    private static OfficeColor ApplyCoverage(OfficeColor color, double coverage) {
        if (coverage >= 0.999D) {
            return color;
        }

        byte alpha = (byte)Math.Round(color.A * Clamp(coverage, 0D, 1D));
        return OfficeColor.FromRgba(color.R, color.G, color.B, alpha);
    }

    private static OfficeColor Interpolate(OfficeColor start, OfficeColor end, double ratio) {
        double inverse = 1D - ratio;
        double alpha = (start.A * inverse) + (end.A * ratio);
        if (alpha <= double.Epsilon) return OfficeColor.Transparent;
        byte r = ToByte(((start.R * start.A * inverse) + (end.R * end.A * ratio)) / alpha);
        byte g = ToByte(((start.G * start.A * inverse) + (end.G * end.A * ratio)) / alpha);
        byte b = ToByte(((start.B * start.A * inverse) + (end.B * end.A * ratio)) / alpha);
        byte a = ToByte(alpha);
        return OfficeColor.FromRgba(r, g, b, a);
    }

    private static byte ToByte(double value) =>
        (byte)Math.Max(0, Math.Min(255, (int)Math.Round(value)));

    private static OfficeColor InterpolateGradient(OfficeLinearGradient gradient, double ratio) {
        return InterpolateGradientStops(gradient.Stops, ratio);
    }

    private static OfficeColor InterpolateGradient(OfficeRadialGradient gradient, double ratio) {
        return InterpolateGradientStops(gradient.Stops, ratio);
    }

    private static OfficeColor InterpolateGradientStops(IReadOnlyList<OfficeGradientStop> stops, double ratio) {
        if (ratio <= stops[0].Offset) {
            return stops[0].Color;
        }

        for (int i = 1; i < stops.Count; i++) {
            OfficeGradientStop next = stops[i];
            if (ratio <= next.Offset) {
                OfficeGradientStop previous = stops[i - 1];
                double span = next.Offset - previous.Offset;
                double localRatio = span <= double.Epsilon ? 0D : (ratio - previous.Offset) / span;
                return Interpolate(previous.Color, next.Color, Clamp(localRatio, 0D, 1D));
            }
        }

        return stops[stops.Count - 1].Color;
    }

    private static double ComputeRadialRatio(OfficeRadialGradient gradient, double x, double y) {
        double endRadiusX = Math.Max(gradient.EndRadiusX, 0.0000001D);
        double endRadiusY = Math.Max(gradient.EndRadiusY, 0.0000001D);
        double normalizedX = (x - gradient.EndX) / endRadiusX;
        double normalizedY = (y - gradient.EndY) / endRadiusY;
        double startX = (gradient.StartX - gradient.EndX) / endRadiusX;
        double startY = (gradient.StartY - gradient.EndY) / endRadiusY;
        double startRadius = gradient.StartRadiusX / endRadiusX;
        double vx = normalizedX - startX;
        double vy = normalizedY - startY;
        double dx = -startX;
        double dy = -startY;
        double dr = 1D - startRadius;
        double a = (dx * dx) + (dy * dy) - (dr * dr);
        double b = -2D * ((vx * dx) + (vy * dy) + (startRadius * dr));
        double c = (vx * vx) + (vy * vy) - (startRadius * startRadius);
        if (Math.Abs(a) < 0.0000001D) {
            if (Math.Abs(b) < 0.0000001D) {
                return 0D;
            }

            return Clamp(-c / b, 0D, 1D);
        }

        double discriminant = (b * b) - (4D * a * c);
        if (discriminant < 0D) {
            return 0D;
        }

        double sqrt = Math.Sqrt(discriminant);
        double t1 = (-b - sqrt) / (2D * a);
        double t2 = (-b + sqrt) / (2D * a);
        double ratio = Math.Max(t1, t2);
        if (ratio < 0D) {
            ratio = Math.Min(t1, t2);
        }

        return Clamp(ratio, 0D, 1D);
    }

    private static byte InterpolateByte(byte start, byte end, double ratio) =>
        (byte)Math.Max(0, Math.Min(255, (int)Math.Round(start + ((end - start) * ratio))));

    private static int Clamp(int value, int min, int max) => value < min ? min : value > max ? max : value;

    private static double Clamp(double value, double min, double max) => value < min ? min : value > max ? max : value;
}
