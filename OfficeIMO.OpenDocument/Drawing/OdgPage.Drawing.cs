using OfficeIMO.Drawing;
using System.Threading;

namespace OfficeIMO.OpenDocument;

public sealed partial class OdgPage {
    /// <summary>Projects supported backgrounds, master artwork and page shapes, applying transforms and layer visibility.</summary>
    /// <param name="lossPolicy">Treatment of unsupported content and approximations.</param>
    /// <param name="forPrint">Selects print-visible layers instead of screen-visible layers.</param>
    public OdfConversionResult<OfficeDrawing> ToDrawing(OdfConversionLossPolicy lossPolicy = OdfConversionLossPolicy.ReportOnly, bool forPrint = false) =>
        ToDrawing(lossPolicy, forPrint, CancellationToken.None);

    /// <summary>Projects supported shapes with cooperative cancellation before each shape and nested group.</summary>
    public OdfConversionResult<OfficeDrawing> ToDrawing(OdfConversionLossPolicy lossPolicy, bool forPrint, CancellationToken cancellationToken) =>
        ToDrawing(lossPolicy, forPrint, cancellationToken, null);

    /// <summary>Projects a page with deterministic, opt-in date/time field evaluation.</summary>
    public OdfConversionResult<OfficeDrawing> ToDrawing(OdfConversionLossPolicy lossPolicy, bool forPrint,
        CancellationToken cancellationToken, OdfDateTimeFieldProjectionOptions? dateTimeFields) =>
        ToDrawing(lossPolicy, forPrint, cancellationToken, 0, 0, new OdfDateTimeFieldProjection(_document, dateTimeFields));

    /// <summary>Projects a page using caller-supplied fonts and shaping before assessing text fitting and clipping.</summary>
    /// <remarks>The returned drawing retains the profile's resources. Changing those resources later can change text layout without updating this operation's loss report.</remarks>
    public OdfConversionResult<OfficeDrawing> ToDrawing(OfficeRenderingProfile renderingProfile,
        OdfConversionLossPolicy lossPolicy = OdfConversionLossPolicy.ReportOnly, bool forPrint = false,
        CancellationToken cancellationToken = default, OdfDateTimeFieldProjectionOptions? dateTimeFields = null) {
        if (renderingProfile == null) throw new ArgumentNullException(nameof(renderingProfile));
        return ToDrawing(lossPolicy, forPrint, cancellationToken, 0, 0, new OdfDateTimeFieldProjection(_document, dateTimeFields), renderingProfile, null);
    }

    // The document projection already knows each page's position in its operation-local snapshot.
    internal OdfConversionResult<OfficeDrawing> ToDrawing(OdfConversionLossPolicy lossPolicy, bool forPrint,
        CancellationToken cancellationToken, int logicalPageNumber, int logicalPageCount, OdfDateTimeFieldProjection dateTimeFields) =>
        ToDrawing(lossPolicy, forPrint, cancellationToken, logicalPageNumber, logicalPageCount, dateTimeFields, null, null);

    internal OdfConversionResult<OfficeDrawing> ToDrawing(OdfConversionLossPolicy lossPolicy, bool forPrint,
        CancellationToken cancellationToken, int logicalPageNumber, int logicalPageCount, OdfDateTimeFieldProjection dateTimeFields,
        OfficeRenderingProfile? renderingProfile, OfficeDrawingTextMetrics? layoutMetrics) {
        cancellationToken.ThrowIfCancellationRequested();
        var report = new OdfConversionReport("ODG", "OfficeDrawing");
        var drawing = new OfficeDrawing(Width.ToPoints(), Height.ToPoints());
        ApplyDrawingProfile(drawing, renderingProfile);
        ProjectBackground(drawing, report);
        var fields = new DrawingFieldContext(this, logicalPageNumber, logicalPageCount, dateTimeFields, cancellationToken);
        if (Master is XElement textMaster)
            PrepareTextFrames(new OdgShapes(_document, textMaster), drawing, report, EffectiveLayers, forPrint, cancellationToken, fields, layoutMetrics);
        PrepareTextFrames(Shapes, drawing, report, EffectiveLayers, forPrint, cancellationToken, fields, layoutMetrics);
        if (Master is XElement master) {
            ProjectShapes(new OdgShapes(_document, master), drawing, report, EffectiveLayers, forPrint, OfficeTransform.Identity, cancellationToken, fields, layoutMetrics);
        }
        ProjectShapes(Shapes, drawing, report, EffectiveLayers, forPrint, OfficeTransform.Identity, cancellationToken, fields, layoutMetrics);
        return new OdfConversionResult<OfficeDrawing>(drawing, report).ApplyPolicy(lossPolicy);
    }

    private static void ProjectShapes(OdgShapes shapes, OfficeDrawing drawing, OdfConversionReport report, OdgLayers layers, bool forPrint, OfficeTransform parentTransform, CancellationToken cancellationToken, DrawingFieldContext fields, OfficeDrawingTextMetrics? layoutMetrics) {
        foreach (OdgShape shape in shapes) {
            cancellationToken.ThrowIfCancellationRequested();
            string feature = "shape:" + shape.Name;
            try {
                if (shape.Layer is string layerName) {
                    OdgLayer? layer = layers.Find(layerName);
                    if (layer == null) report.Add("layer:" + layerName, OdfConversionMappingStatus.Approximated, message: "Undefined layer is treated as visible.");
                    else if (!layer.IsVisible(forPrint)) continue;
                }
                OfficeTransform transform = OdfDrawingTransform.Parse(shape.Transform).Then(parentTransform);
                if (shape.IsGroup) {
                    if (shape.HasOpacityGradient || shape.FillOpacity < 1D || shape.StrokeOpacity < 1D)
                        report.Add(feature + ":group-opacity", OdfConversionMappingStatus.Unsupported,
                            message: "Group paint opacity is preserved in ODF but is not applied to child shapes in this projection.");
                    ProjectShapes(shape.Children, drawing, report, layers, forPrint, transform, cancellationToken, fields, layoutMetrics); continue;
                }
                double x, y, w, h;
                OfficeShape? geometry = null;
                IReadOnlyList<OfficePathCommand>? linePath = null;
                if (shape.ElementName is "line" or "connector") {
                    OdgShape routeShape = shape.IsConnector ? shape.WithProjectionBounds(fields.ProjectedFrameBounds) : shape;
                    if (shape.IsConnector && (shape.ConnectorKind != OdgConnectorKind.Line || shape.Element.Attribute(OdfNamespaces.Svg + "d") != null)) {
                        IReadOnlyList<OfficePathCommand> path = routeShape.ConnectorRouteCommands;
                        linePath = path;
                        var positions = OdgShape.ConnectorPathPoints(path).ToArray();
                        x = positions.Min(point => point.X); y = positions.Min(point => point.Y);
                        w = positions.Max(point => point.X) - x; h = positions.Max(point => point.Y) - y;
                        geometry = OfficeShape.Path(Math.Max(w, 0.001), Math.Max(h, 0.001), path.Select(command => command.Translate(x, y)));
                    } else {
                        double x1 = routeShape.X1.ToPoints(), y1 = routeShape.Y1.ToPoints(), x2 = routeShape.X2.ToPoints(), y2 = routeShape.Y2.ToPoints();
                        x = Math.Min(x1, x2); y = Math.Min(y1, y2); w = Math.Abs(x2 - x1); h = Math.Abs(y2 - y1);
                        geometry = OfficeShape.Line(x1 - x, y1 - y, x2 - x, y2 - y);
                        linePath = new[] { OfficePathCommand.MoveTo(x1, y1), OfficePathCommand.LineTo(x2, y2) };
                    }
                    if (shape.UsesAutomaticGlue) report.Add(feature + ":automatic-glue", OdfConversionMappingStatus.Approximated,
                        message: "Automatic attachment retains a saved edge center when valid, otherwise selects the nearest edge center; native obstacle routing is not recalculated.");
                } else if (shape.ElementName == "custom-shape") {
                    OdfRect bounds = shape.Bounds;
                    x = bounds.X.ToPoints(); y = bounds.Y.ToPoints(); w = bounds.Width.ToPoints(); h = bounds.Height.ToPoints();
                    geometry = ProjectEnhancedGeometry(shape, w, h, report, feature);
                } else if (shape.ElementName is "rect" or "ellipse" or "circle" || shape.HasLocalGeometry || shape.IsImage || shape.Element.Element(OdfNamespaces.Draw + "text-box") != null) {
                    OdfRect bounds = shape.Bounds;
                    x = bounds.X.ToPoints(); y = bounds.Y.ToPoints(); w = bounds.Width.ToPoints(); h = bounds.Height.ToPoints();
                    if (!shape.IsImage) geometry = shape.CreateDrawingGeometry(w, h);
                } else if (OdfTextTraversal.Paragraphs(shape.TextRoot).Any()) {
                    OdfRect bounds = shape.Bounds;
                    x = bounds.X.ToPoints(); y = bounds.Y.ToPoints(); w = bounds.Width.ToPoints(); h = bounds.Height.ToPoints();
                    report.Add(feature, OdfConversionMappingStatus.Skipped, message: "Native " + shape.ElementName + " geometry is preserved in ODG but not projected; supported text is projected separately.");
                } else {
                    report.Add(feature, OdfConversionMappingStatus.Skipped, message: "Native " + shape.ElementName + " is preserved in ODG but not projected."); continue;
                }
                var local = new OfficeDrawing(Math.Max(w, 0.001), Math.Max(h, 0.001));
                CopyDrawingResources(drawing, local);
                OfficeDrawing? preparedText = null;
                if (fields.PreparedTextFrames.TryGetValue(shape.Element, out TextFrameProjection textProjection)) {
                    preparedText = textProjection.Drawing;
                    w = textProjection.Width; h = textProjection.Height;
                    local = new OfficeDrawing(preparedText.Width, preparedText.Height);
                    CopyDrawingResources(preparedText, local);
                    // A zero-area ordinary frame has no rectangle paint. Keep the
                    // semantic size for clipping without inventing a tiny shape or
                    // weakening the shared positioned-shape size contract.
                    geometry = w == 0D || h == 0D ? null : shape.CreateDrawingGeometry(w, h);
                }
                bool imageProjected = false;
                ReportUnselectedFrameText(shape, report, feature);
                if (shape.IsImage) {
                    byte[]? data = shape.GetImageBytes();
                    if (!TryReadProjectedImage(data, cancellationToken, out OfficeImageInfo image, out int svgLosses)) {
                        report.Add(feature, OdfConversionMappingStatus.Skipped, message: "Image bytes are unavailable or outside the renderable profile; external references are not fetched.");
                    } else {
                        if (svgLosses > 0) report.Add(feature + ":svg-content", OdfConversionMappingStatus.Unsupported,
                            message: "The shared SVG renderer omits or approximates " + svgLosses.ToString(CultureInfo.InvariantCulture) + " unsupported declarations; source bytes are preserved.");
                        OfficeImageProjection? projection = ProjectImageCrop(shape, image, w, h, report);
                        projection = ProjectImageMirror(shape, projection, w, h, report);
                        if (projection.HasValue) {
                            local.AddImage(data!, image.MimeType, projection.Value, opacity: shape.ImageOpacity ?? 1D);
                            imageProjected = true;
                        }
                    }
                } else if (geometry != null) {
                    ApplyPaint(shape, geometry!, report, transform);
                    if (shape.ElementName == "custom-shape") geometry.FillRule = OfficeFillRule.EvenOdd;
                    if (shape.ElementName is "line" or "connector") geometry!.FillColor = null;
                    if (!ProjectStrokeContours(shape, geometry!, local, report)) local.AddShape(geometry!, 0, 0);
                }
                if (preparedText != null) local.AddDrawing(preparedText, 0, 0);
                else if (shape.ElementName is not ("line" or "connector")) ProjectText(shape, local, report, w, h, fields, cancellationToken: cancellationToken, layoutMetrics: layoutMetrics);
                if (geometry != null && (transform != OfficeTransform.Identity || x < 0 || y < 0 || x + local.Width > drawing.Width || y + local.Height > drawing.Height)) {
                    local = ExpandStrokeCanvas(local, geometry, out double left, out double top);
                    x -= left; y -= top;
                }
                if (transform == OfficeTransform.Identity && x >= 0 && y >= 0 && x + local.Width <= drawing.Width && y + local.Height <= drawing.Height)
                    drawing.AddDrawing(local, x, y);
                else drawing.AddEffectDrawing(local, OfficeTransform.Translate(x, y).Then(transform));
                if (geometry != null || imageProjected)
                    report.Add(feature, OdfConversionMappingStatus.Approximated, message: "Geometry, transforms, supported fill, solid or dashed stroke, caps and joins with uniform opacity and supported image cropping and horizontal mirroring are projected; advanced graphic styles, effects and metadata are not reproduced.");
                if (linePath != null) ProjectLineText(shape, drawing, report, linePath, transform, fields, cancellationToken, layoutMetrics);
            } catch (Exception exception) when (exception is FormatException || exception is ArgumentException || exception is NotSupportedException || exception is InvalidDataException || exception is OverflowException) {
                report.Add(feature, OdfConversionMappingStatus.Skipped, message: exception.Message);
            }
        }
    }
    private static void ApplyPaint(OdgShape source, OfficeShape target, OdfConversionReport report, OfficeTransform transform) {
        target.FillColor = ToColor(source.FillColor); target.StrokeColor = ToColor(source.StrokeColor); target.StrokeWidth = source.StrokeWidth?.ToPoints() ?? 1;
        target.FillOpacity = ReadFillOpacity(source, report); target.StrokeOpacity = source.StrokeOpacity;
        if (target.StrokeColor.HasValue && target.StrokeWidth == 0) {
            target.StrokeWidth = 0.25;
            report.Add("hairline-stroke", OdfConversionMappingStatus.Approximated, message: "A device-dependent zero-width stroke is projected as a 0.25 point line.");
        }
        target.FillRule = source.FillRule;
        ApplyGradientFill(source, target, report, transform);
        if (target.StrokeColor.HasValue) {
            OdfStrokeDashPattern? dash = source.ResolveStrokeDash();
            target.StrokeLineCap = source.StrokeLineCap ?? (dash?.RoundCaps == true ? OfficeStrokeLineCap.Round : OfficeStrokeLineCap.Butt);
            target.StrokeLineJoin = source.StrokeLineJoin ?? OfficeStrokeLineJoin.Round;
            if (dash != null) target.SetStrokeDashArray(dash.Resolve(target.StrokeWidth));
            if (source.HasStackedStrokeDashes) report.Add("shape:" + source.Name + ":stacked-dashes", OdfConversionMappingStatus.Unsupported,
                message: "Additional overlaid stroke-dash definitions are preserved but only the primary stroke is projected.");
        }
    }
    private static double? ReadFillOpacity(OdgShape shape, OdfConversionReport report) {
        if (!shape.HasOpacityGradient) return shape.FillOpacity;
        report.Add("shape:" + shape.Name + ":opacity-gradient", OdfConversionMappingStatus.Unsupported,
            message: "Opacity gradients are preserved in ODF but projected with an opaque fill.");
        return null;
    }
    private static OfficeColor? ToColor(OdfColor? color) => color.HasValue ? OfficeColor.Parse(color.Value.ToString()) : (OfficeColor?)null;
}
