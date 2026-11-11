using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

internal sealed partial class VisioDrawingProjection {
    private void ProjectShape(VisioPage page, OfficeDrawing drawing, VisioShape shape,
        VisioRenderProjection projection, VisioNativeTextStyleResolver styles,
        List<OfficeImageExportDiagnostic> imageDiagnostics, string location) {
        VisioNativeShapeTransform transform;
        try {
            transform = VisioNativeShapeTransform.Create(shape, imageDiagnostics, location);
        } catch (Exception exception) when (exception is ArgumentException || exception is System.IO.InvalidDataException) {
            _report.Add("VISIO_SHAPE_TRANSFORM_INVALID", exception.Message, OfficeConversionLossKind.Omission, location);
            return;
        }
        try {
            if (VisioForeignImage.IsForeign(shape)) {
                if (VisioForeignImage.TryGetProjection(shape, page, projection.GeometryDensity, imageDiagnostics, location, out OfficeImageProjection image)) {
                    OfficeRasterImage raster = VisioForeignImage.Decode(shape, null, imageDiagnostics, location, _token);
                    ChargeImagePixels(raster);
                    byte[] bytes = OfficePngWriter.Encode(raster, _token);
                    ChargeImageBytes(bytes);
                    drawing.AddClippedImageWithInterpolation(bytes, "image/png", image.Translate(0, projection.ContentOffsetY), true,
                        0, 0, OfficeClipPath.Rectangle(drawing.Width, drawing.Height));
                }
            } else {
                ProjectGeometry(drawing, shape, projection, transform, location);
            }
            if (!string.IsNullOrEmpty(shape.Text)) {
                VisioTextFramePlacement frame = VisioTextFramePlacement.Resolve(shape, projection.DrawingToPhysical, transform);
                (double x, double y) = projection.PagePoint(frame.PageX, frame.PageY);
                ProjectText(drawing, shape.Text!, VisioTextBackground.Effective(shape),
                    VisioRichTextProjection.Create(page, shape, 72D, _token, styles),
                    x, y, frame.ContentWidth * projection.GeometryDensity, frame.ContentHeight * projection.GeometryDensity,
                    frame.Angle, 10D, location);
            }
            ProjectLinks(drawing, shape.Hyperlinks, ShapeCorners(shape, projection, transform), location);
            ReportShapeData(shape.ShapeData.Count + shape.Data.Count, location);
            if (shape.PreservedShapeChildren.Any(entry => entry.RawElement != null && entry.RawElement.Name.LocalName != "ForeignData"
                && !IsProjectedSection(entry.RawElement)))
                _report.Add("VISIO_DRAWING_SHAPE_CONTENT", "Opaque shape children are retained in the source but are not projected.", OfficeConversionLossKind.Omission, location);
        } catch (ArgumentException exception) {
            _report.Add("VISIO_DRAWING_SHAPE_INVALID", exception.Message, OfficeConversionLossKind.Omission, location);
        }
    }

    // Cached gradient stops are painted; other opaque children remain reported.
    private static bool IsProjectedSection(System.Xml.Linq.XElement element) =>
        element.Name.LocalName == "Section" && (string?)element.Attribute("N") == "FillGradient";

    private void ProjectGeometry(OfficeDrawing drawing, VisioShape shape, VisioRenderProjection projection, VisioNativeShapeTransform transform, string location) {
        bool native = VisioShapeGeometry.TryGetRenderClosedPaths(shape, out List<VisioShapeGeometryPath> paths);
        if (!native) {
            string kind = VisioShapeGeometry.ResolveRenderKind(shape);
            paths.Add(kind is "ellipse" or "circle"
                ? VisioShapeGeometry.CreateBuiltinEllipsePath(shape)
                : new VisioShapeGeometryPath(VisioShapeGeometry.GetBuiltinClosedPath(shape, kind), false, false, true, 0));
            _report.Add("VISIO_DRAWING_GEOMETRY_FALLBACK", "A named builtin outline or rectangular fallback replaces unavailable native geometry.", OfficeConversionLossKind.Approximation, location);
        }
        if (paths.Any(path => path.HasFlattenedCurves))
            _report.Add("VISIO_DRAWING_CURVES", "Spline and NURBS rows are flattened by the shared Visio geometry owner; lines, arcs, ellipses and Bezier rows remain exact curves.", OfficeConversionLossKind.Approximation, location);
        OfficePoint Map(double x, double y) {
            OfficePoint page = transform.PagePoint(x, y);
            (double px, double py) = projection.PagePoint(page.X, page.Y);
            return new OfficePoint(px, py);
        }
        var rendered = new List<(VisioShapeGeometryPath Path, List<OfficePathCommand> Commands)>();
        foreach (VisioShapeGeometryPath path in paths) {
            ChargePoints(path.Points.Count);
            rendered.Add((path, CreateCommands(path, Map, path.IsClosed)));
        }
        bool fill = shape.FillPattern != 0 && shape.FillColor.A > 0;
        bool stroke = shape.LinePattern != 0 && shape.LineWeight > 0 && shape.LineColor.A > 0;
        FillPaint paint = fill ? ResolveFill(shape, Map, location) : default;
        foreach (var group in rendered.GroupBy(item => item.Path.FillGroup)) {
            if (fill) {
                var commands = new List<OfficePathCommand>();
                foreach (var item in group.Where(item => item.Path.CanFill)) AppendFillContour(commands, item.Path, item.Commands);
                if (commands.Count > 0) AddPath(drawing, commands, paint, null, 0, OfficeStrokeDashStyle.Solid);
            }
            if (stroke) foreach (var item in group.Where(item => !item.Path.NoLine && item.Path.Points.Count >= 2)) {
                AddPath(drawing, item.Commands, default, shape.LineColor, shape.LineWeight * projection.PhysicalDensity,
                    OfficeStrokeDashStyleMapper.FromVisioLinePattern(shape.LinePattern));
            }
        }
    }

    private FillPaint ResolveFill(VisioShape shape, Func<double, double, OfficePoint> map, string location) {
        if (VisioFillGradient.TryResolve(shape, out VisioFillGradient? gradient, out bool unreadable)) {
            double width = shape.WidthInches, height = shape.HeightInches;
            if (gradient!.IsLinear) {
                // FillGradientAngle follows the Office convention: clockwise from left-to-right as displayed,
                // so the local y-up direction negates the sine. The line spans the frame's projected extent.
                double dx = Math.Cos(gradient.Angle), dy = -Math.Sin(gradient.Angle);
                double half = (Math.Abs(width * dx) + Math.Abs(height * dy)) / 2D;
                OfficePoint start = map(width / 2D - dx * half, height / 2D - dy * half);
                OfficePoint end = map(width / 2D + dx * half, height / 2D + dy * half);
                if (start.X != end.X || start.Y != end.Y)
                    return new FillPaint(null, OfficeLinearGradient.CreateImported(start.X, start.Y, end.X, end.Y, gradient.Stops), null, default);
            } else {
                // The unit circle of the gradient frame reaches the corners of the shape's local frame.
                double reach = Math.Sqrt(2D);
                OfficePoint origin = map(width / 2D - reach * width / 2D, height / 2D + reach * height / 2D);
                OfficePoint unitX = map(width / 2D + reach * width / 2D, height / 2D + reach * height / 2D);
                OfficePoint unitY = map(width / 2D - reach * width / 2D, height / 2D - reach * height / 2D);
                var frame = new OfficeTransform(unitX.X - origin.X, unitX.Y - origin.Y, unitY.X - origin.X, unitY.Y - origin.Y, origin.X, origin.Y);
                if (frame.TryInvert(out _)) {
                    if (gradient.Direction != 1)
                        _report.Add("VISIO_DRAWING_FILL_GRADIENT", "The " + VisioFillGradient.Describe(gradient.Direction) + " gradient is drawn as a centred radial gradient.", OfficeConversionLossKind.Approximation, location);
                    return new FillPaint(null, null, new OfficeRadialGradient(0.5, 0.5, 0, 0.5, 0.5, 0.5, gradient.Stops), frame);
                }
            }
        } else if (unreadable) {
            _report.Add("VISIO_DRAWING_FILL_GRADIENT", "The cached foreground color replaces a gradient whose stops are not cached literal colors.", OfficeConversionLossKind.Approximation, location);
        }
        if (shape.FillPattern != 1)
            _report.Add("VISIO_DRAWING_FILL_PATTERN", "The cached foreground color replaces a non-solid fill pattern.", OfficeConversionLossKind.Approximation, location);
        return new FillPaint(shape.FillColor, null, null, default);
    }

    /// <summary>Builds move, line and cubic commands; flattened points are used only where no exact segment exists.</summary>
    private static List<OfficePathCommand> CreateCommands(VisioShapeGeometryPath path, Func<double, double, OfficePoint> map, bool close) {
        var commands = new List<OfficePathCommand>();
        if (path.Segments is { Count: > 1 } segments) {
            commands.Add(OfficePathCommand.MoveTo(map(segments[0].End.X, segments[0].End.Y)));
            for (int index = 1; index < segments.Count; index++) {
                VisioGeometrySegment segment = segments[index];
                commands.Add(segment.IsCubic
                    ? OfficePathCommand.CubicBezierTo(map(segment.Control1.X, segment.Control1.Y), map(segment.Control2.X, segment.Control2.Y), map(segment.End.X, segment.End.Y))
                    : OfficePathCommand.LineTo(map(segment.End.X, segment.End.Y)));
            }
        } else if (path.Points.Count >= 2) {
            commands.Add(OfficePathCommand.MoveTo(map(path.Points[0].X, path.Points[0].Y)));
            for (int index = 1; index < path.Points.Count; index++) commands.Add(OfficePathCommand.LineTo(map(path.Points[index].X, path.Points[index].Y)));
        }
        if (close && commands.Count > 0) commands.Add(OfficePathCommand.Close());
        return commands;
    }

    // Fill closes open contours implicitly, as the native fill does.
    private static void AppendFillContour(List<OfficePathCommand> commands, VisioShapeGeometryPath path, List<OfficePathCommand> outline) {
        if (outline.Count == 0) return;
        commands.AddRange(outline);
        if (!path.IsClosed) commands.Add(OfficePathCommand.Close());
    }

    private static void AppendPath(List<OfficePathCommand> commands, IReadOnlyList<OfficePoint> points, bool close) {
        if (points.Count < 2) return;
        commands.Add(OfficePathCommand.MoveTo(points[0].X, points[0].Y));
        for (int index = 1; index < points.Count; index++) commands.Add(OfficePathCommand.LineTo(points[index].X, points[index].Y));
        if (close) commands.Add(OfficePathCommand.Close());
    }

    private static IEnumerable<OfficePoint> CommandPoints(IEnumerable<OfficePathCommand> commands) {
        foreach (OfficePathCommand command in commands) {
            if (command.Kind == OfficePathCommandKind.Close) continue;
            if (command.Kind == OfficePathCommandKind.CubicBezierTo) { yield return command.ControlPoint1; yield return command.ControlPoint2; }
            else if (command.Kind == OfficePathCommandKind.QuadraticBezierTo) yield return command.ControlPoint1;
            yield return command.Point;
        }
    }

    private static void AddPath(OfficeDrawing drawing, List<OfficePathCommand> commands, FillPaint fill,
        OfficeColor? stroke, double weight, OfficeStrokeDashStyle dash, bool rounded = false) {
        if (commands.Count == 0) return;
        // Keep off-page coordinates in the placement transform so each path has valid, non-negative local
        // bounds. A curve lies inside its control polygon, so control points bound it too.
        OfficePoint[] points = CommandPoints(commands).ToArray();
        bool overflow = points.Any(point => point.X < 0 || point.Y < 0 || point.X > drawing.Width || point.Y > drawing.Height);
        double x = overflow ? points.Min(point => point.X) : 0D;
        double y = overflow ? points.Min(point => point.Y) : 0D;
        double width = overflow ? Math.Max(1D, points.Max(point => point.X) - x) : drawing.Width;
        double height = overflow ? Math.Max(1D, points.Max(point => point.Y) - y) : drawing.Height;
        OfficeShape geometry = OfficeShape.Path(width, height, overflow ? commands.Select(command => command.Translate(x, y)) : commands);
        geometry.FillColor = fill.Color; geometry.StrokeColor = stroke; geometry.StrokeWidth = weight;
        if (fill.Linear != null || fill.Radial != null) {
            // Gradients are defined in page coordinates; express them in the path's normalized local box.
            OfficeTransform pageToBox = OfficeTransform.Translate(-x, -y).Then(OfficeTransform.Scale(1D / width, 1D / height));
            geometry.FillGradient = fill.Linear?.TransformCoordinates(pageToBox);
            geometry.FillRadialGradient = fill.Radial?.TransformCoordinates(fill.RadialFrame.Then(pageToBox));
        }
        geometry.StrokeDashStyle = dash; geometry.FillRule = OfficeFillRule.EvenOdd;
        if (rounded) { geometry.StrokeLineCap = OfficeStrokeLineCap.Round; geometry.StrokeLineJoin = OfficeStrokeLineJoin.Round; }
        if (overflow) {
            // Raster effect groups have local image bounds. Include the stroke's paint
            // envelope so translating a route cannot crop its visible half at an edge.
            double padding = stroke.HasValue ? weight * (rounded ? 1D : geometry.StrokeMiterLimit) : 0D;
            var local = new OfficeDrawing(width + padding * 2D, height + padding * 2D);
            local.AddShape(geometry, padding, padding);
            drawing.AddEffectDrawing(local, OfficeTransform.Translate(x - padding, y - padding));
        } else drawing.AddShape(geometry, 0, 0);
    }

    private static IReadOnlyList<OfficePoint> ShapeCorners(VisioShape shape, VisioRenderProjection projection, VisioNativeShapeTransform transform) {
        var corners = new List<OfficePoint>(4);
        foreach ((double X, double Y) local in new[] { (0D, 0D), (shape.WidthInches, 0D), (shape.WidthInches, shape.HeightInches), (0D, shape.HeightInches) }) {
            OfficePoint page = transform.PagePoint(local.X, local.Y);
            (double x, double y) = projection.PagePoint(page.X, page.Y);
            corners.Add(new OfficePoint(x, y));
        }
        return corners;
    }

    private void ReportShapeData(int dataCount, string location) {
        if (dataCount > 0)
            _report.Add("VISIO_DRAWING_METADATA", "Shape data is retained in the source but is not included in diagram page content.", OfficeConversionLossKind.Omission, location);
    }

    /// <summary>Fill paint in drawing coordinates: a solid color, a linear gradient, or a radial gradient with its unit frame.</summary>
    private readonly struct FillPaint {
        internal FillPaint(OfficeColor? color, OfficeLinearGradient? linear, OfficeRadialGradient? radial, OfficeTransform radialFrame) {
            Color = color; Linear = linear; Radial = radial; RadialFrame = radialFrame;
        }

        internal OfficeColor? Color { get; }

        internal OfficeLinearGradient? Linear { get; }

        internal OfficeRadialGradient? Radial { get; }

        internal OfficeTransform RadialFrame { get; }
    }
}
