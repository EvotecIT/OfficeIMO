namespace OfficeIMO.Drawing;

/// <summary>
/// A soft shadow or glow as one silhouette of the shape blurred with a Gaussian kernel. Raster, SVG and PDF writers
/// share it so the effect keeps one geometry, radius and opacity across outputs.
/// </summary>
internal sealed class OfficeBlurredShapeEffect {
    // Office blur and glow radii describe where the effect fades out. A shadow fades over its blur radius on both
    // sides of the edge (two standard deviations each way). A glow's outline grows by half its radius and fades out
    // over the other half (three standard deviations each way), so it stays at full opacity on the contour.
    private const double ShadowSigmaPerBlurRadius = 0.5D;
    private const double GlowSigmaPerRadius = 1D / 6D;
    private const double MarginSigmas = 3D;
    // Two pixels per standard deviation, interpolated, already match an exact Gaussian within a few levels.
    private const double MaximumPixelsPerSigma = 2D;
    private const double MinimumPixelDeviation = 0.05D;

    /// <summary>Pixel budget for one blurred effect; larger effects render at a lower density.</summary>
    internal const long MaximumRasterPixels = 4L * 1024L * 1024L;

    private OfficeBlurredShapeEffect(OfficeShape silhouette, OfficeColor color, double opacity, double sigma,
        double offsetX, double offsetY, OfficeTransform? transform) {
        Silhouette = silhouette;
        Color = color;
        Opacity = opacity;
        Sigma = sigma;
        OffsetX = offsetX;
        OffsetY = offsetY;
        Transform = transform;
        ClipPath = silhouette.ClipPath;
        silhouette.ClipPath = null;
        double strokeReach = silhouette.StrokeColor.HasValue ? Math.Max(0D, silhouette.StrokeWidth) / 2D : 0D;
        // An unspecified join paints as a miter.
        if (silhouette.StrokeColor.HasValue && silhouette.StrokeLineJoin is null or OfficeStrokeLineJoin.Miter) strokeReach *= Math.Max(1D, silhouette.StrokeMiterLimit);
        Margin = Math.Ceiling(strokeReach + sigma * MarginSigmas + 1D);
        // Paths, polygons and lines can reach outside their declared box; the blur surface follows the geometry.
        double left = 0D, top = 0D, right = silhouette.Width, bottom = silhouette.Height;
        if (silhouette.Kind is OfficeShapeKind.Path or OfficeShapeKind.Polygon or OfficeShapeKind.Line) {
            foreach (OfficeFlattenedPathContour contour in OfficeStrokeGeometry.FlattenShape(silhouette, 1D)) {
                foreach (OfficePoint point in contour.Points) {
                    left = Math.Min(left, point.X); top = Math.Min(top, point.Y);
                    right = Math.Max(right, point.X); bottom = Math.Max(bottom, point.Y);
                }
            }
        }
        LocalLeft = Math.Floor(left) - Margin;
        LocalTop = Math.Floor(top) - Margin;
        LocalWidth = Math.Ceiling(right) + Margin - LocalLeft;
        LocalHeight = Math.Ceiling(bottom) + Margin - LocalTop;
    }

    /// <summary>Shape-local geometry painted in <see cref="Color"/>; it has no transform, clip, effects or gradients.</summary>
    internal OfficeShape Silhouette { get; }

    /// <summary>Opaque effect color; <see cref="Opacity"/> carries the effect and color alpha.</summary>
    internal OfficeColor Color { get; }

    internal double Opacity { get; }

    /// <summary>Gaussian standard deviation in shape-local units.</summary>
    internal double Sigma { get; }

    /// <summary>Offset of the effect in drawing units, applied after the shape transform.</summary>
    internal double OffsetX { get; }

    internal double OffsetY { get; }

    internal OfficeTransform? Transform { get; }

    /// <summary>The shape's local clip. Writers apply it after the blur, as it bounds everything the shape paints.</summary>
    internal OfficeClipPath? ClipPath { get; }

    /// <summary>Local distance around the silhouette geometry that holds the visible blur.</summary>
    internal double Margin { get; }

    /// <summary>Shape-local box of the blurred silhouette, including its margin.</summary>
    internal double LocalLeft { get; }

    internal double LocalTop { get; }

    internal double LocalWidth { get; }

    internal double LocalHeight { get; }

    /// <summary>Returns the blurred shadow, or <see langword="null"/> for none or a hard (unblurred) shadow.</summary>
    internal static OfficeBlurredShapeEffect? CreateShadow(OfficeShape shape) {
        OfficeShadow? shadow = shape.Shadow;
        if (shadow == null || shadow.Opacity <= 0D || shadow.Color.A == 0 || !(shadow.BlurRadius > 0D)) return null;
        if (shape.ClipPath?.Kind == OfficeClipPathKind.Empty) return null;
        ResolveShadowPaint(shape, out bool hasFill, out bool hasStroke);
        OfficeColor color = OfficeColor.FromRgb(shadow.Color.R, shadow.Color.G, shadow.Color.B);
        OfficeShape silhouette = CreateSilhouette(shape, color, paintFill: hasFill || !hasStroke, paintStroke: hasStroke,
            Math.Max(0D, shape.StrokeWidth), shape.StrokeLineCap, shape.StrokeLineJoin);
        double opacity = Math.Min(1D, shadow.Opacity) * shadow.Color.A / 255D;
        return new OfficeBlurredShapeEffect(silhouette, color, opacity, shadow.BlurRadius * ShadowSigmaPerBlurRadius,
            shadow.OffsetX, shadow.OffsetY, shape.Transform);
    }

    /// <summary>Returns the glow halo, or <see langword="null"/> when the shape has none.</summary>
    internal static OfficeBlurredShapeEffect? CreateGlow(OfficeShape shape) {
        OfficeGlow? glow = shape.Glow;
        if (glow == null || !(glow.Radius > 0D) || glow.Opacity <= 0D || glow.Color.A == 0) return null;
        if (shape.ClipPath?.Kind == OfficeClipPathKind.Empty) return null;
        OfficeColor color = OfficeColor.FromRgb(glow.Color.R, glow.Color.G, glow.Color.B);
        double opacity = Math.Min(1D, glow.Opacity) * glow.Color.A / 255D;
        double sigma = glow.Radius * GlowSigmaPerRadius;
        double growth = (Math.Max(0D, shape.StrokeWidth) + glow.Radius) / 2D;
        if (OfficeShadowLayerPlanner.CanExpand(shape) && HasOpaqueSolidFill(shape)) {
            // An opaque fill hides the inside of the halo, so the grown shape filled is the same silhouette and
            // far cheaper to rasterize than a wide round-joined outline.
            OfficeShape expanded = OfficeShadowLayerPlanner.CreateExpandedShape(shape, growth);
            if (shape.Kind == OfficeShapeKind.Rectangle) {
                // Grown corners stay round, as a round-joined outline would draw them.
                OfficeShape rounded = OfficeShape.RoundedRectangle(expanded.Width, expanded.Height, growth);
                rounded.FillColor = color;
                expanded = rounded;
            }
            OfficeShape grown = CreateSilhouette(expanded, color,
                paintFill: true, paintStroke: false, 0D, shape.StrokeLineCap, shape.StrokeLineJoin);
            return new OfficeBlurredShapeEffect(grown, color, opacity, sigma, -growth, -growth, transform: null);
        }
        // Round corners whatever the source join, as Office draws a glow.
        OfficeShape silhouette = CreateSilhouette(shape, color, paintFill: false, paintStroke: true,
            growth * 2D, OfficeStrokeLineCap.Round, OfficeStrokeLineJoin.Round);
        return new OfficeBlurredShapeEffect(silhouette, color, opacity, sigma, 0D, 0D, shape.Transform);
    }

    private static bool HasOpaqueSolidFill(OfficeShape shape) =>
        shape.Kind != OfficeShapeKind.Line && shape.FillGradient == null && shape.FillRadialGradient == null
        && shape.FillColor is { A: 255 } && shape.FillOpacity.GetValueOrDefault(1D) >= 1D;

    internal static void ResolveShadowPaint(OfficeShape shape, out bool hasFill, out bool hasStroke) {
        hasStroke = shape.Kind == OfficeShapeKind.Line ||
            (shape.StrokeWidth > 0D &&
                (shape.StrokeRadialGradient != null ||
                 shape.StrokeGradient != null ||
                 (shape.StrokeColor.HasValue && shape.StrokeColor.Value.A > 0)));
        hasFill = shape.Kind != OfficeShapeKind.Line &&
            (shape.FillRadialGradient != null || shape.FillGradient != null || (shape.FillColor.HasValue && shape.FillColor.Value.A > 0));
    }

    private static OfficeShape CreateSilhouette(OfficeShape shape, OfficeColor color, bool paintFill, bool paintStroke,
        double strokeWidth, OfficeStrokeLineCap? lineCap, OfficeStrokeLineJoin? lineJoin) {
        OfficeShape silhouette = shape.Clone();
        silhouette.Shadow = null;
        silhouette.Glow = null;
        silhouette.Transform = null;
        silhouette.FillGradient = null;
        silhouette.FillRadialGradient = null;
        silhouette.FillColor = paintFill && shape.Kind != OfficeShapeKind.Line ? color : null;
        silhouette.FillOpacity = 1D;
        silhouette.StrokeGradient = null;
        silhouette.StrokeRadialGradient = null;
        silhouette.StrokeColor = paintStroke && strokeWidth > 0D ? color : null;
        silhouette.StrokeOpacity = 1D;
        silhouette.StrokeWidth = paintStroke ? strokeWidth : 0D;
        silhouette.StrokeDashStyle = OfficeStrokeDashStyle.Solid;
        silhouette.SetStrokeDashArray(null);
        silhouette.StrokeHairline = false;
        silhouette.StrokeStartMarker = null;
        silhouette.StrokeEndMarker = null;
        silhouette.StrokeLineCap = lineCap;
        silhouette.StrokeLineJoin = lineJoin;
        return silhouette;
    }

    /// <summary>Returns the larger axis scale of a shape transform, one for none.</summary>
    internal static double GetAxisScale(OfficeTransform? transform) {
        if (!transform.HasValue) return 1D;
        OfficeTransform value = transform.Value;
        return Math.Max(Math.Sqrt(value.M11 * value.M11 + value.M12 * value.M12),
            Math.Sqrt(value.M21 * value.M21 + value.M22 * value.M22));
    }

    /// <summary>
    /// Chooses the raster density for the blurred silhouette: at most the output density, at most two pixels per
    /// standard deviation, and within the pixel budget.
    /// </summary>
    internal double ChoosePixelsPerUnit(double outputPixelsPerUnit, long maximumPixels) {
        double density = Math.Min(outputPixelsPerUnit, MaximumPixelsPerSigma / Sigma);
        double area = LocalWidth * LocalHeight;
        if (area > 0D && density * density * area > maximumPixels) density = Math.Sqrt(maximumPixels / area);
        return Math.Max(density, 1E-6D);
    }

    /// <summary>
    /// Renders the blurred silhouette coverage, one byte per pixel, whose pixels cover <see cref="LocalLeft"/>,
    /// <see cref="LocalTop"/>, <see cref="LocalWidth"/> and <see cref="LocalHeight"/> in shape-local units.
    /// <see cref="Opacity"/> is left to the caller.
    /// </summary>
    /// <param name="pixelsPerUnit">Raster density in pixels per shape-local unit.</param>
    /// <param name="cancellationToken">Cancels rendering and filtering.</param>
    /// <param name="width">Coverage width in pixels.</param>
    /// <param name="height">Coverage height in pixels.</param>
    /// <param name="defaultLineCap">Cap for a silhouette whose source leaves it unspecified, so the blur follows the
    /// writer's own stroke default (PDF uses butt caps; the raster renderer rounds them).</param>
    /// <param name="defaultLineJoin">Join for a silhouette whose source leaves it unspecified.</param>
    internal byte[] RenderCoverage(double pixelsPerUnit, CancellationToken cancellationToken, out int width, out int height,
        OfficeStrokeLineCap? defaultLineCap = null, OfficeStrokeLineJoin? defaultLineJoin = null) {
        OfficeShape silhouette = Silhouette;
        if (silhouette.StrokeColor.HasValue && (silhouette.StrokeLineCap == null && defaultLineCap != null
            || silhouette.StrokeLineJoin == null && defaultLineJoin != null)) {
            silhouette = silhouette.Clone();
            silhouette.StrokeLineCap ??= defaultLineCap;
            silhouette.StrokeLineJoin ??= defaultLineJoin;
        }
        string? key = CreateCoverageKey(silhouette, pixelsPerUnit);
        if (key != null && CoverageCache.TryGet(key, out byte[] cached, out width, out height)) return cached;
        byte[] coverage = RenderCoverageCore(silhouette, pixelsPerUnit, cancellationToken, out width, out height);
        if (key != null) CoverageCache.Add(key, coverage, width, height);
        return coverage;
    }

    // Themes repeat one effect on many equal shapes, so primitive silhouettes share their blurred coverage.
    // Paths and polygons are rendered each time rather than keyed by their geometry.
    private string? CreateCoverageKey(OfficeShape s, double pixelsPerUnit) {
        if (s.Kind != OfficeShapeKind.Rectangle && s.Kind != OfficeShapeKind.RoundedRectangle && s.Kind != OfficeShapeKind.Ellipse) return null;
        return string.Join("|", ((int)s.Kind).ToString(System.Globalization.CultureInfo.InvariantCulture),
            Format(s.Width), Format(s.Height), Format(s.CornerRadius), s.FillColor.HasValue ? "f" : "-",
            s.StrokeColor.HasValue ? Format(s.StrokeWidth) : "-", s.StrokeLineCap?.ToString() ?? "-",
            s.StrokeLineJoin?.ToString() ?? "-", Format(s.StrokeMiterLimit),
            Format(Sigma), Format(Margin), Format(pixelsPerUnit));
    }

    private static string Format(double value) => value.ToString("R", System.Globalization.CultureInfo.InvariantCulture);

    private static class CoverageCache {
        private const long MaximumBytes = 16L * 1024L * 1024L;
        private static readonly object Sync = new();
        private static readonly Dictionary<string, (byte[] Coverage, int Width, int Height)> Entries = new(StringComparer.Ordinal);
        private static long _bytes;

        internal static bool TryGet(string key, out byte[] coverage, out int width, out int height) {
            lock (Sync) {
                if (Entries.TryGetValue(key, out var entry)) {
                    (coverage, width, height) = entry;
                    return true;
                }
            }
            coverage = Array.Empty<byte>(); width = height = 0;
            return false;
        }

        // Callers never write to a returned plane, so one array can serve every hit.
        internal static void Add(string key, byte[] coverage, int width, int height) {
            if (coverage.LongLength > MaximumBytes / 4) return;
            lock (Sync) {
                if (Entries.ContainsKey(key)) return;
                if (_bytes + coverage.LongLength > MaximumBytes) { Entries.Clear(); _bytes = 0; }
                Entries.Add(key, (coverage, width, height));
                _bytes += coverage.LongLength;
            }
        }
    }

    private byte[] RenderCoverageCore(OfficeShape shape, double pixelsPerUnit, CancellationToken cancellationToken, out int width, out int height) {
        int pixelWidth = Math.Max(1, (int)Math.Ceiling(LocalWidth * pixelsPerUnit));
        int pixelHeight = Math.Max(1, (int)Math.Ceiling(LocalHeight * pixelsPerUnit));
        var drawing = new OfficeDrawing(pixelWidth / pixelsPerUnit, pixelHeight / pixelsPerUnit).AddShape(shape, -LocalLeft, -LocalTop);
        OfficeRasterImage silhouette = OfficeDrawingRasterRenderer.Render(drawing, new OfficeDrawingRasterRenderOptions {
            Scale = pixelsPerUnit,
            CancellationToken = cancellationToken
        });
        width = silhouette.Width;
        height = silhouette.Height;
        byte[] pixels = silhouette.PixelBuffer;
        var coverage = new byte[width * height];
        for (int i = 0; i < coverage.Length; i++) coverage[i] = pixels[i * 4 + 3];
        double deviation = Sigma * pixelsPerUnit;
        // A deviation far below a pixel leaves the antialiased silhouette unchanged.
        return deviation < MinimumPixelDeviation ? coverage
            : OfficeRasterFilters.GaussianBlurCoverage(coverage, width, height, deviation, cancellationToken);
    }

    /// <summary>Renders the blurred silhouette as <see cref="Color"/> with the blur as alpha.</summary>
    internal OfficeRasterImage Render(double pixelsPerUnit, CancellationToken cancellationToken) {
        byte[] coverage = RenderCoverage(pixelsPerUnit, cancellationToken, out int width, out int height);
        var image = new OfficeRasterImage(width, height);
        byte[] pixels = image.PixelBuffer;
        for (int i = 0, offset = 0; i < coverage.Length; i++, offset += 4) {
            pixels[offset] = Color.R; pixels[offset + 1] = Color.G; pixels[offset + 2] = Color.B; pixels[offset + 3] = coverage[i];
        }
        return image;
    }

    /// <summary>Maps pixels of an image from <see cref="Render"/> to drawing units: local box, shape transform, offset.</summary>
    internal OfficeTransform CreateImageTransform(double pixelsPerUnit, double shapeX, double shapeY) =>
        OfficeTransform.Scale(1D / pixelsPerUnit, 1D / pixelsPerUnit)
            .Then(OfficeTransform.Translate(LocalLeft, LocalTop))
            .Then(Transform ?? OfficeTransform.Identity)
            .Then(OfficeTransform.Translate(shapeX + OffsetX, shapeY + OffsetY));
}
