using System;
using System.Collections.Generic;

namespace OfficeIMO.Drawing;

/// <summary>
/// Dependency-free drawing canvas for an <see cref="OfficeRasterImage"/>.
/// </summary>
public sealed partial class OfficeRasterCanvas {
    private const int AntiAliasSamples = 3;
    private const double MinimumDashSegmentAdvance = 1E-9D;
    private const double MinimumRasterDashLength = 0.25D;
    private static readonly OfficeTrueTypeFont? DefaultFont = OfficeTrueTypeFont.TryLoadDefault();
    private readonly OfficeRasterImage? _image;
    private readonly OfficeRasterRenderTarget? _target;
    private readonly OfficeTrueTypeFont? _font;
    private readonly OfficeFontFaceCollection? _fonts;
    private readonly IOfficeTextShapingProvider? _textShapingProvider;
    private readonly string? _textShapingLanguage;
    private readonly ICollection<OfficeImageExportDiagnostic>? _diagnosticSink;
    private readonly string? _diagnosticSource;
    private readonly System.Threading.CancellationToken _cancellationToken;
    private bool _reportedBoundedTextShapingFallback;
    private bool _reportedIncompleteTextShapingFallback;
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
        _font = font ?? DefaultFont;
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
        _font = font ?? DefaultFont;
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

    internal System.Threading.CancellationToken CancellationToken => _cancellationToken;

    internal ICollection<OfficeImageExportDiagnostic>? DiagnosticSink => _diagnosticSink;

    internal string? DiagnosticSource => _diagnosticSource;

    /// <summary>Fills a rectangle.</summary>
    public void FillRectangle(double x, double y, double width, double height, OfficeColor color) {
        if (color.A == 0 || width <= 0D || height <= 0D || IsOutsideCanvas(x, y, width, height)) {
            return;
        }

        int left = Clamp((int)Math.Floor(x), 0, Width - 1);
        int top = Clamp((int)Math.Floor(y), 0, Height - 1);
        int right = Clamp((int)Math.Ceiling(x + width), 0, Width);
        int bottom = Clamp((int)Math.Ceiling(y + height), 0, Height);
        for (int py = top; py < bottom; py++) {
            for (int px = left; px < right; px++) {
                BlendPixel(px, py, color);
            }
        }
    }

    /// <summary>Fills a rectangle with a linear gradient.</summary>
    public void FillLinearGradientRectangle(double x, double y, double width, double height, OfficeLinearGradient gradient) {
        if (gradient == null) {
            throw new ArgumentNullException(nameof(gradient));
        }

        if (width <= 0D || height <= 0D || IsOutsideCanvas(x, y, width, height)) {
            return;
        }

        OfficeGradientStop start = gradient.Stops[0];
        int left = Clamp((int)Math.Floor(x), 0, Width - 1);
        int top = Clamp((int)Math.Floor(y), 0, Height - 1);
        int right = Clamp((int)Math.Ceiling(x + width), 0, Width);
        int bottom = Clamp((int)Math.Ceiling(y + height), 0, Height);
        double dx = gradient.EndX - gradient.StartX;
        double dy = gradient.EndY - gradient.StartY;
        double lengthSquared = (dx * dx) + (dy * dy);
        if (lengthSquared <= double.Epsilon) {
            FillRectangle(x, y, width, height, start.Color);
            return;
        }

        for (int py = top; py < bottom; py++) {
            double ny = ((py + 0.5D) - y) / height;
            for (int px = left; px < right; px++) {
                double nx = ((px + 0.5D) - x) / width;
                double ratio = (((nx - gradient.StartX) * dx) + ((ny - gradient.StartY) * dy)) / lengthSquared;
                BlendPixel(px, py, InterpolateGradient(gradient, Clamp(ratio, 0D, 1D)));
            }
        }
    }

    /// <summary>Fills a rectangle with a radial gradient.</summary>
    public void FillRadialGradientRectangle(double x, double y, double width, double height, OfficeRadialGradient gradient) {
        if (gradient == null) {
            throw new ArgumentNullException(nameof(gradient));
        }

        if (width <= 0D || height <= 0D || IsOutsideCanvas(x, y, width, height)) {
            return;
        }

        int left = Clamp((int)Math.Floor(x), 0, Width - 1);
        int top = Clamp((int)Math.Floor(y), 0, Height - 1);
        int right = Clamp((int)Math.Ceiling(x + width), 0, Width);
        int bottom = Clamp((int)Math.Ceiling(y + height), 0, Height);
        for (int py = top; py < bottom; py++) {
            double ny = ((py + 0.5D) - y) / height;
            for (int px = left; px < right; px++) {
                double nx = ((px + 0.5D) - x) / width;
                BlendPixel(px, py, InterpolateGradient(gradient, ComputeRadialRatio(gradient, nx, ny)));
            }
        }
    }

    /// <summary>Draws a rectangle outline.</summary>
    public void DrawRectangle(double x, double y, double width, double height, OfficeColor color, double thickness = 1D) {
        DrawLine(x, y, x + width, y, color, thickness);
        DrawLine(x + width, y, x + width, y + height, color, thickness);
        DrawLine(x + width, y + height, x, y + height, color, thickness);
        DrawLine(x, y + height, x, y, color, thickness);
    }

    /// <summary>Fills an ellipse bounded by the supplied rectangle.</summary>
    public void FillEllipse(double x, double y, double width, double height, OfficeColor color) {
        if (color.A == 0 || width <= 0D || height <= 0D || IsOutsideCanvas(x, y, width, height)) {
            return;
        }

        int left = Clamp((int)Math.Floor(x), 0, Width - 1);
        int top = Clamp((int)Math.Floor(y), 0, Height - 1);
        int right = Clamp((int)Math.Ceiling(x + width), 0, Width - 1);
        int bottom = Clamp((int)Math.Ceiling(y + height), 0, Height - 1);
        for (int py = top; py <= bottom; py++) {
            for (int px = left; px <= right; px++) {
                double coverage = EllipseFillCoverage(px, py, x, y, width, height);
                if (coverage > 0D) {
                    BlendPixel(px, py, ApplyCoverage(color, coverage));
                }
            }
        }
    }

    /// <summary>Draws an ellipse outline bounded by the supplied rectangle.</summary>
    public void DrawEllipse(double x, double y, double width, double height, OfficeColor color, double thickness = 1D) {
        if (color.A == 0 || width <= 0D || height <= 0D || thickness <= 0D || IsOutsideCanvas(x, y, width, height)) {
            return;
        }

        double rx = width / 2D;
        double ry = height / 2D;
        double cx = x + rx;
        double cy = y + ry;
        double innerRx = Math.Max(0.0001D, rx - (thickness / 2D));
        double innerRy = Math.Max(0.0001D, ry - (thickness / 2D));
        double outerRx = rx + (thickness / 2D);
        double outerRy = ry + (thickness / 2D);
        int left = Clamp((int)Math.Floor(cx - outerRx), 0, Width - 1);
        int top = Clamp((int)Math.Floor(cy - outerRy), 0, Height - 1);
        int right = Clamp((int)Math.Ceiling(cx + outerRx), 0, Width - 1);
        int bottom = Clamp((int)Math.Ceiling(cy + outerRy), 0, Height - 1);
        for (int py = top; py <= bottom; py++) {
            for (int px = left; px <= right; px++) {
                double coverage = EllipseStrokeCoverage(px, py, cx, cy, outerRx, outerRy, innerRx, innerRy);
                if (coverage > 0D) {
                    BlendPixel(px, py, ApplyCoverage(color, coverage));
                }
            }
        }
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
        if (radiusX <= 0D || radiusY <= 0D || (fill.A == 0 && (stroke.A == 0 || thickness <= 0D))) {
            return;
        }

        double strokeHalf = Math.Max(0D, thickness / 2D);
        double outerRadiusX = Math.Max(radiusX + strokeHalf, 0.0001D);
        double outerRadiusY = Math.Max(radiusY + strokeHalf, 0.0001D);
        double rotationRadians = OfficeGeometry.DegreesToRadians(rotationDegrees);
        OfficePoint topLeft = OfficeGeometry.RotatePoint(new OfficePoint(centerX - outerRadiusX, centerY - outerRadiusY), rotationCenterX, rotationCenterY, rotationRadians);
        OfficePoint topRight = OfficeGeometry.RotatePoint(new OfficePoint(centerX + outerRadiusX, centerY - outerRadiusY), rotationCenterX, rotationCenterY, rotationRadians);
        OfficePoint bottomRight = OfficeGeometry.RotatePoint(new OfficePoint(centerX + outerRadiusX, centerY + outerRadiusY), rotationCenterX, rotationCenterY, rotationRadians);
        OfficePoint bottomLeft = OfficeGeometry.RotatePoint(new OfficePoint(centerX - outerRadiusX, centerY + outerRadiusY), rotationCenterX, rotationCenterY, rotationRadians);
        double minX = Math.Min(Math.Min(topLeft.X, topRight.X), Math.Min(bottomRight.X, bottomLeft.X));
        double maxX = Math.Max(Math.Max(topLeft.X, topRight.X), Math.Max(bottomRight.X, bottomLeft.X));
        double minY = Math.Min(Math.Min(topLeft.Y, topRight.Y), Math.Min(bottomRight.Y, bottomLeft.Y));
        double maxY = Math.Max(Math.Max(topLeft.Y, topRight.Y), Math.Max(bottomRight.Y, bottomLeft.Y));
        int left = Clamp((int)Math.Floor(minX - 1D), 0, Width - 1);
        int right = Clamp((int)Math.Ceiling(maxX + 1D), 0, Width - 1);
        int top = Clamp((int)Math.Floor(minY - 1D), 0, Height - 1);
        int bottom = Clamp((int)Math.Ceiling(maxY + 1D), 0, Height - 1);
        for (int py = top; py <= bottom; py++) {
            for (int px = left; px <= right; px++) {
                (double FillCoverage, double StrokeCoverage) coverage = RotatedEllipseCoverage(
                    px,
                    py,
                    centerX,
                    centerY,
                    radiusX,
                    radiusY,
                    strokeHalf,
                    fill.A > 0,
                    rotationRadians,
                    rotationCenterX,
                    rotationCenterY);
                if (fill.A > 0 && coverage.FillCoverage > 0D) {
                    BlendPixel(px, py, ApplyCoverage(fill, coverage.FillCoverage));
                }

                if (stroke.A > 0 && thickness > 0D && coverage.StrokeCoverage > 0D) {
                    BlendPixel(px, py, ApplyCoverage(stroke, coverage.StrokeCoverage));
                }
            }
        }
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
        if (color.A == 0 || thickness <= 0D || dashLength <= 0D || gapLength < 0D
            || !IsFinite(dashLength) || !IsFinite(gapLength)) {
            return;
        }

        double length = Distance(x1, y1, x2, y2);
        if (!IsFinite(length) || length <= 0D) {
            return;
        }
        NormalizeRasterDashLengths(ref dashLength, ref gapLength);
        if (!TryClipLineToCanvas(ref x1, ref y1, ref x2, ref y2, thickness, length, out double leadingDistance, out _)) return;
        double cycle = dashLength + gapLength;
        double patternPosition = AdvancePatternPosition(0D, leadingDistance, cycle);
        DrawDashedPathSegment(new OfficePoint(x1, y1), new OfficePoint(x2, y2), color, thickness, dashLength, gapLength, ref patternPosition);
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
        if (color.A == 0 || thickness <= 0D) {
            return;
        }

        if (dashPattern == null || dashPattern.Count == 0) {
            DrawLine(x1, y1, x2, y2, color, thickness);
            return;
        }

        double length = Distance(x1, y1, x2, y2);
        if (!IsFinite(length) || length <= 0D) {
            return;
        }

        List<double> pattern = new List<double>(dashPattern.Count);
        for (int i = 0; i < dashPattern.Count; i++) {
            if (dashPattern[i] > 0D) {
                pattern.Add(dashPattern[i]);
            }
        }

        if (pattern.Count == 0) {
            DrawLine(x1, y1, x2, y2, color, thickness);
            return;
        }
        if ((pattern.Count & 1) == 1) {
            int originalCount = pattern.Count;
            for (int index = 0; index < originalCount; index++) pattern.Add(pattern[index]);
        }
        IReadOnlyList<double> rasterPattern = NormalizeRasterDashPattern(pattern);
        if (!TryClipLineToCanvas(ref x1, ref y1, ref x2, ref y2, thickness, length, out double leadingDistance, out _)) return;
        double cycle = 0D;
        for (int index = 0; index < rasterPattern.Count; index++) cycle += rasterPattern[index];
        double patternPosition = AdvancePatternPosition(0D, leadingDistance, cycle);
        DrawPatternedPathSegment(new OfficePoint(x1, y1), new OfficePoint(x2, y2), color, thickness, rasterPattern, ref patternPosition);
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
        if (color.A == 0 || thickness <= 0D || radiusX <= 0D || radiusY <= 0D) {
            return;
        }

        double sweep = endDegrees - startDegrees;
        if (Math.Abs(sweep) <= 0.0001D) {
            return;
        }

        int segments = Math.Max(4, (int)Math.Ceiling(Math.Abs(sweep) / 10D));
        double startRadians = OfficeGeometry.DegreesToRadians(startDegrees);
        double sweepRadians = OfficeGeometry.DegreesToRadians(sweep);
        double rotationRadians = OfficeGeometry.DegreesToRadians(rotationDegrees);
        OfficePoint previous = CreateArcStartPoint(centerX, centerY, radiusX, radiusY, startRadians, rotationRadians, rotationCenterX, rotationCenterY);
        foreach (OfficePoint current in OfficeGeometry.CreateEllipticalArcPoints(
            centerX,
            centerY,
            radiusX,
            radiusY,
            startRadians,
            sweepRadians,
            segments,
            rotationRadians,
            rotationCenterX,
            rotationCenterY)) {
            DrawLine(previous.X, previous.Y, current.X, current.Y, color, thickness);
            previous = current;
        }
    }

    private void DrawLineSegment(double x1, double y1, double x2, double y2, OfficeColor color, double thickness) {
        double radius = Math.Max(0.5D, thickness / 2D);
        double outer = radius + 1D;
        int left = Clamp((int)Math.Floor(Math.Min(x1, x2) - outer), 0, Width - 1);
        int top = Clamp((int)Math.Floor(Math.Min(y1, y2) - outer), 0, Height - 1);
        int right = Clamp((int)Math.Ceiling(Math.Max(x1, x2) + outer), 0, Width - 1);
        int bottom = Clamp((int)Math.Ceiling(Math.Max(y1, y2) + outer), 0, Height - 1);
        for (int py = top; py <= bottom; py++) {
            for (int px = left; px <= right; px++) {
                double distance = DistanceToSegment(px + 0.5D, py + 0.5D, x1, y1, x2, y2);
                double coverage = Clamp(radius + 0.5D - distance, 0D, 1D);
                if (coverage > 0D) {
                    BlendPixel(px, py, ApplyCoverage(color, coverage));
                }
            }
        }
    }

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

        for (int i = 1; i < points.Count; i++) {
            DrawLine(points[i - 1].X, points[i - 1].Y, points[i].X, points[i].Y, color, thickness);
        }

        if (points.Count > 2) {
            DrawLine(points[points.Count - 1].X, points[points.Count - 1].Y, points[0].X, points[0].Y, color, thickness);
        }
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
            projection.FlipVertical);
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
        OfficeTransform imageTransform = projection.CreateUnitSquareTransform();
        if (!imageTransform.TryInvert(out OfficeTransform inverseTransform)) {
            return;
        }

        (double minX, double minY, double maxX, double maxY) = projection.GetDestinationBounds();
        int left = Clamp((int)Math.Floor(minX), 0, Width - 1);
        int top = Clamp((int)Math.Floor(minY), 0, Height - 1);
        int right = Clamp((int)Math.Ceiling(maxX), 0, Width - 1);
        int bottom = Clamp((int)Math.Ceiling(maxY), 0, Height - 1);
        bool cropped = sourceLeft > 0D || sourceTop > 0D || sourceWidth < 1D || sourceHeight < 1D;
        for (int py = top; py <= bottom; py++) {
            for (int px = left; px <= right; px++) {
                OfficePoint unit = inverseTransform.TransformPoint(new OfficePoint(px + 0.5D, py + 0.5D));
                double u = unit.X;
                double v = unit.Y;
                if (u < 0D || u >= 1D || v < 0D || v >= 1D) {
                    continue;
                }

                double sourceX;
                double sourceY;
                if (cropped) {
                    sourceX = (sourceLeft * image.Width) + (u * Math.Max(0D, (sourceWidth * image.Width) - 1D));
                    sourceY = (sourceTop * image.Height) + (v * Math.Max(0D, (sourceHeight * image.Height) - 1D));
                } else {
                    sourceX = (u * image.Width) - 0.5D;
                    sourceY = (v * image.Height) - 0.5D;
                }

                BlendPixel(px, py, SampleBilinear(image, sourceX, sourceY));
            }
        }
    }

    /// <summary>Draws an image through an arbitrary destination-space affine transform.</summary>
    public void DrawAffineImage(OfficeRasterImage image, OfficeTransform transform, double opacity = 1D) {
        if (image == null) throw new ArgumentNullException(nameof(image));
        if (double.IsNaN(opacity) || double.IsInfinity(opacity) || opacity < 0D || opacity > 1D) {
            throw new ArgumentOutOfRangeException(nameof(opacity), "Image opacity must be between zero and one.");
        }
        if (opacity <= 0D || !transform.TryInvert(out OfficeTransform inverse)) return;

        (double minX, double minY, double maxX, double maxY) = transform.TransformRectangleBounds(0D, 0D, image.Width, image.Height);
        int left = Clamp((int)Math.Floor(minX), 0, Width - 1);
        int top = Clamp((int)Math.Floor(minY), 0, Height - 1);
        int right = Clamp((int)Math.Ceiling(maxX), 0, Width - 1);
        int bottom = Clamp((int)Math.Ceiling(maxY), 0, Height - 1);
        for (int py = top; py <= bottom; py++) {
            for (int px = left; px <= right; px++) {
                OfficePoint source = inverse.TransformPoint(new OfficePoint(px + 0.5D, py + 0.5D));
                if (source.X < 0D || source.X >= image.Width || source.Y < 0D || source.Y >= image.Height) continue;
                OfficeColor color = SampleBilinear(image, source.X - 0.5D, source.Y - 0.5D);
                if (opacity < 1D) color = OfficeColor.FromRgba(color.R, color.G, color.B, (byte)Math.Round(color.A * opacity));
                BlendPixel(px, py, color);
            }
        }
    }

    private void FillContours(IReadOnlyList<IReadOnlyList<OfficePoint>> contours, OfficeColor color, OfficeFillRule fillRule) {
        if (color.A == 0 || contours == null || contours.Count == 0) {
            return;
        }

        bool found = false;
        double minX = 0D;
        double maxX = 0D;
        double minY = 0D;
        double maxY = 0D;
        for (int i = 0; i < contours.Count; i++) {
            IReadOnlyList<OfficePoint> contour = contours[i];
            if (contour.Count < 3) {
                continue;
            }

            for (int j = 0; j < contour.Count; j++) {
                OfficePoint point = contour[j];
                if (!found) {
                    minX = maxX = point.X;
                    minY = maxY = point.Y;
                    found = true;
                } else {
                    minX = Math.Min(minX, point.X);
                    maxX = Math.Max(maxX, point.X);
                    minY = Math.Min(minY, point.Y);
                    maxY = Math.Max(maxY, point.Y);
                }
            }
        }

        if (!found) {
            return;
        }

        int left = Clamp((int)Math.Floor(minX), 0, Width - 1);
        int right = Clamp((int)Math.Ceiling(maxX), 0, Width - 1);
        int top = Clamp((int)Math.Floor(minY), 0, Height - 1);
        int bottom = Clamp((int)Math.Ceiling(maxY), 0, Height - 1);
        for (int py = top; py <= bottom; py++) {
            for (int px = left; px <= right; px++) {
                double coverage = ContoursCoverage(contours, px, py, fillRule);
                if (coverage > 0D) {
                    BlendPixel(px, py, ApplyCoverage(color, coverage));
                }
            }
        }
    }

    private void FillPolygonCore(IReadOnlyList<OfficePoint> points, OfficeColor color) {
        double minX = points[0].X;
        double maxX = points[0].X;
        double minY = points[0].Y;
        double maxY = points[0].Y;
        for (int i = 1; i < points.Count; i++) {
            minX = Math.Min(minX, points[i].X);
            maxX = Math.Max(maxX, points[i].X);
            minY = Math.Min(minY, points[i].Y);
            maxY = Math.Max(maxY, points[i].Y);
        }

        int left = Clamp((int)Math.Floor(minX), 0, Width - 1);
        int right = Clamp((int)Math.Ceiling(maxX), 0, Width - 1);
        int top = Clamp((int)Math.Floor(minY), 0, Height - 1);
        int bottom = Clamp((int)Math.Ceiling(maxY), 0, Height - 1);
        for (int py = top; py <= bottom; py++) {
            for (int px = left; px <= right; px++) {
                double coverage = PolygonCoverage(points, px, py);
                if (coverage > 0D) {
                    BlendPixel(px, py, ApplyCoverage(color, coverage));
                }
            }
        }
    }

    private void FillPolygonCore(IReadOnlyList<OfficePoint> points, OfficeLinearGradient gradient) {
        double minX = points[0].X;
        double maxX = points[0].X;
        double minY = points[0].Y;
        double maxY = points[0].Y;
        for (int i = 1; i < points.Count; i++) {
            minX = Math.Min(minX, points[i].X);
            maxX = Math.Max(maxX, points[i].X);
            minY = Math.Min(minY, points[i].Y);
            maxY = Math.Max(maxY, points[i].Y);
        }

        double width = Math.Max(0.0001D, maxX - minX);
        double height = Math.Max(0.0001D, maxY - minY);
        OfficeGradientStop start = gradient.Stops[0];
        double dx = gradient.EndX - gradient.StartX;
        double dy = gradient.EndY - gradient.StartY;
        double lengthSquared = (dx * dx) + (dy * dy);
        if (lengthSquared <= double.Epsilon) {
            FillPolygonCore(points, start.Color);
            return;
        }

        int left = Clamp((int)Math.Floor(minX), 0, Width - 1);
        int right = Clamp((int)Math.Ceiling(maxX), 0, Width - 1);
        int top = Clamp((int)Math.Floor(minY), 0, Height - 1);
        int bottom = Clamp((int)Math.Ceiling(maxY), 0, Height - 1);
        for (int py = top; py <= bottom; py++) {
            double ny = ((py + 0.5D) - minY) / height;
            for (int px = left; px <= right; px++) {
                double coverage = PolygonCoverage(points, px, py);
                if (coverage <= 0D) {
                    continue;
                }

                double nx = ((px + 0.5D) - minX) / width;
                double ratio = (((nx - gradient.StartX) * dx) + ((ny - gradient.StartY) * dy)) / lengthSquared;
                BlendPixel(px, py, ApplyCoverage(InterpolateGradient(gradient, Clamp(ratio, 0D, 1D)), coverage));
            }
        }
    }

    private void FillPolygonCore(IReadOnlyList<OfficePoint> points, OfficeRadialGradient gradient) {
        double minX = points[0].X;
        double maxX = points[0].X;
        double minY = points[0].Y;
        double maxY = points[0].Y;
        for (int i = 1; i < points.Count; i++) {
            minX = Math.Min(minX, points[i].X);
            maxX = Math.Max(maxX, points[i].X);
            minY = Math.Min(minY, points[i].Y);
            maxY = Math.Max(maxY, points[i].Y);
        }

        double width = Math.Max(0.0001D, maxX - minX);
        double height = Math.Max(0.0001D, maxY - minY);
        int left = Clamp((int)Math.Floor(minX), 0, Width - 1);
        int right = Clamp((int)Math.Ceiling(maxX), 0, Width - 1);
        int top = Clamp((int)Math.Floor(minY), 0, Height - 1);
        int bottom = Clamp((int)Math.Ceiling(maxY), 0, Height - 1);
        for (int py = top; py <= bottom; py++) {
            double ny = ((py + 0.5D) - minY) / height;
            for (int px = left; px <= right; px++) {
                double coverage = PolygonCoverage(points, px, py);
                if (coverage <= 0D) {
                    continue;
                }

                double nx = ((px + 0.5D) - minX) / width;
                BlendPixel(px, py, ApplyCoverage(InterpolateGradient(gradient, ComputeRadialRatio(gradient, nx, ny)), coverage));
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

    private double PolygonCoverage(IReadOnlyList<OfficePoint> points, int x, int y) {
        int samples = CoverageSamples;
        int covered = 0;
        for (int sy = 0; sy < samples; sy++) {
            double sampleY = y + (sy + 0.5D) / samples;
            for (int sx = 0; sx < samples; sx++) {
                double sampleX = x + (sx + 0.5D) / samples;
                if (ContainsPoint(points, sampleX, sampleY)) {
                    covered++;
                }
            }
        }

        return covered / (double)(samples * samples);
    }

    private double ContoursCoverage(IReadOnlyList<IReadOnlyList<OfficePoint>> contours, int x, int y, OfficeFillRule fillRule) {
        int samples = CoverageSamples;
        int covered = 0;
        for (int sy = 0; sy < samples; sy++) {
            double sampleY = y + (sy + 0.5D) / samples;
            for (int sx = 0; sx < samples; sx++) {
                double sampleX = x + (sx + 0.5D) / samples;
                int winding = 0;
                for (int i = 0; i < contours.Count; i++) {
                    IReadOnlyList<OfficePoint> contour = contours[i];
                    if (contour.Count < 3) {
                        continue;
                    }

                    if (fillRule == OfficeFillRule.NonZero) {
                        winding += GetWindingNumber(contour, sampleX, sampleY);
                    } else if (ContainsPoint(contour, sampleX, sampleY)) {
                        winding++;
                    }
                }

                if (fillRule == OfficeFillRule.NonZero ? winding != 0 : (winding & 1) == 1) {
                    covered++;
                }
            }
        }

        return covered / (double)(samples * samples);
    }

    private double EllipseFillCoverage(int x, int y, double ellipseX, double ellipseY, double width, double height) {
        double rx = width / 2D;
        double ry = height / 2D;
        double cx = ellipseX + rx;
        double cy = ellipseY + ry;
        int samples = CoverageSamples;
        int covered = 0;
        for (int sy = 0; sy < samples; sy++) {
            double sampleY = y + (sy + 0.5D) / samples;
            for (int sx = 0; sx < samples; sx++) {
                double sampleX = x + (sx + 0.5D) / samples;
                double dx = (sampleX - cx) / Math.Max(rx, 0.0001D);
                double dy = (sampleY - cy) / Math.Max(ry, 0.0001D);
                if ((dx * dx) + (dy * dy) <= 1D) {
                    covered++;
                }
            }
        }

        return covered / (double)(samples * samples);
    }

    private double EllipseStrokeCoverage(int x, int y, double cx, double cy, double outerRx, double outerRy, double innerRx, double innerRy) {
        int samples = CoverageSamples;
        int covered = 0;
        for (int sy = 0; sy < samples; sy++) {
            double sampleY = y + (sy + 0.5D) / samples;
            for (int sx = 0; sx < samples; sx++) {
                double sampleX = x + (sx + 0.5D) / samples;
                double ox = (sampleX - cx) / Math.Max(outerRx, 0.0001D);
                double oy = (sampleY - cy) / Math.Max(outerRy, 0.0001D);
                double ix = (sampleX - cx) / innerRx;
                double iy = (sampleY - cy) / innerRy;
                if ((ox * ox) + (oy * oy) <= 1D && (ix * ix) + (iy * iy) >= 1D) {
                    covered++;
                }
            }
        }

        return covered / (double)(samples * samples);
    }

    private (double FillCoverage, double StrokeCoverage) RotatedEllipseCoverage(
        int x,
        int y,
        double centerX,
        double centerY,
        double radiusX,
        double radiusY,
        double strokeHalf,
        bool hasFill,
        double rotationRadians,
        double rotationCenterX,
        double rotationCenterY) {
        int filled = 0;
        int stroked = 0;
        double outerRadiusX = Math.Max(radiusX + strokeHalf, 0.0001D);
        double outerRadiusY = Math.Max(radiusY + strokeHalf, 0.0001D);
        double innerRadiusX = Math.Max(radiusX - strokeHalf, 0.0001D);
        double innerRadiusY = Math.Max(radiusY - strokeHalf, 0.0001D);
        int samples = CoverageSamples;
        for (int sy = 0; sy < samples; sy++) {
            double sampleY = y + (sy + 0.5D) / samples;
            for (int sx = 0; sx < samples; sx++) {
                double sampleX = x + (sx + 0.5D) / samples;
                OfficePoint local = Math.Abs(rotationRadians) > 0.0001D
                    ? OfficeGeometry.RotatePoint(new OfficePoint(sampleX, sampleY), rotationCenterX, rotationCenterY, -rotationRadians)
                    : new OfficePoint(sampleX, sampleY);
                double dx = local.X - centerX;
                double dy = local.Y - centerY;
                double fillMetric = (dx * dx / (radiusX * radiusX)) + (dy * dy / (radiusY * radiusY));
                if (hasFill && fillMetric <= 1D) {
                    filled++;
                    continue;
                }

                double outerMetric = (dx * dx / (outerRadiusX * outerRadiusX)) + (dy * dy / (outerRadiusY * outerRadiusY));
                double innerMetric = (dx * dx / (innerRadiusX * innerRadiusX)) + (dy * dy / (innerRadiusY * innerRadiusY));
                if (outerMetric <= 1D && innerMetric >= 1D) {
                    stroked++;
                }
            }
        }

        double sampleCount = samples * samples;
        return (filled / sampleCount, stroked / sampleCount);
    }

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

    private static double DistanceToSegment(double px, double py, double x1, double y1, double x2, double y2) {
        double dx = x2 - x1;
        double dy = y2 - y1;
        double lengthSquared = (dx * dx) + (dy * dy);
        if (lengthSquared < 0.0001D) {
            double sx = px - x1;
            double sy = py - y1;
            return Math.Sqrt((sx * sx) + (sy * sy));
        }

        double t = ((px - x1) * dx + (py - y1) * dy) / lengthSquared;
        t = Clamp(t, 0D, 1D);
        double projectionX = x1 + t * dx;
        double projectionY = y1 + t * dy;
        double distanceX = px - projectionX;
        double distanceY = py - projectionY;
        return Math.Sqrt((distanceX * distanceX) + (distanceY * distanceY));
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
        byte r = InterpolateByte(start.R, end.R, ratio);
        byte g = InterpolateByte(start.G, end.G, ratio);
        byte b = InterpolateByte(start.B, end.B, ratio);
        byte a = InterpolateByte(start.A, end.A, ratio);
        return OfficeColor.FromRgba(r, g, b, a);
    }

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
