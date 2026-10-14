using System.Xml.Linq;
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
                ProjectText(drawing, shape.Text!, shape.TextStyle,
                    VisioRichTextProjection.Create(page, shape, 72D, _token, styles),
                    x, y, frame.ContentWidth * projection.GeometryDensity, frame.ContentHeight * projection.GeometryDensity,
                    frame.Angle, 10D, location);
            }
            ReportMetadata(shape.ShapeData.Count + shape.Data.Count, shape.Hyperlinks.Count, location);
            if (shape.PreservedShapeChildren.Any(entry => entry.RawElement != null && entry.RawElement.Name.LocalName != "ForeignData"))
                _report.Add("VISIO_DRAWING_SHAPE_CONTENT", "Opaque shape children are retained in the source but are not projected.", OfficeConversionLossKind.Omission, location);
        } catch (ArgumentException exception) {
            _report.Add("VISIO_DRAWING_SHAPE_INVALID", exception.Message, OfficeConversionLossKind.Omission, location);
        }
    }

    private void ProjectGeometry(OfficeDrawing drawing, VisioShape shape, VisioRenderProjection projection, VisioNativeShapeTransform transform, string location) {
        bool native = VisioShapeGeometry.TryGetRenderClosedPaths(shape, out List<VisioShapeGeometryPath> paths);
        if (!native) {
            string kind = VisioShapeGeometry.ResolveRenderKind(shape);
            List<(double X, double Y)> points = kind is "ellipse" or "circle"
                ? OfficeGeometry.CreateEllipticalArcPointsAsTuples(shape.Width / 2D, shape.Height / 2D,
                    shape.Width / 2D, shape.Height / 2D, 0, Math.PI * 2D, 64).ToList()
                : VisioShapeGeometry.GetBuiltinClosedPath(shape, kind);
            paths.Add(new VisioShapeGeometryPath(points, false, false, true, 0));
            _report.Add("VISIO_DRAWING_GEOMETRY_FALLBACK", "A named builtin outline or rectangular fallback replaces unavailable native geometry.", OfficeConversionLossKind.Approximation, location);
        } else if (HasCurveRows(shape.PreservedGeometrySections.Concat((shape.MasterShape ?? shape.Master?.Shape)?.PreservedGeometrySections ?? Enumerable.Empty<XElement>()))) {
            _report.Add("VISIO_DRAWING_CURVES", "Cached curve rows are flattened by the shared Visio geometry owner.", OfficeConversionLossKind.Approximation, location);
        }
        var rendered = new List<(VisioShapeGeometryPath Path, List<OfficePoint> Points)>();
        foreach (VisioShapeGeometryPath path in paths) {
            ChargePoints(path.Points.Count);
            rendered.Add((path, path.Points.Select(point => {
                OfficePoint page = transform.PagePoint(point.X, point.Y);
                (double x, double y) = projection.PagePoint(page.X, page.Y);
                return new OfficePoint(x, y);
            }).ToList()));
        }
        bool fill = shape.FillPattern != 0 && shape.FillColor.A > 0;
        bool stroke = shape.LinePattern != 0 && shape.LineWeight > 0 && shape.LineColor.A > 0;
        if (fill && shape.FillPattern != 1)
            _report.Add("VISIO_DRAWING_FILL_PATTERN", "The cached foreground color replaces a non-solid fill pattern.", OfficeConversionLossKind.Approximation, location);
        foreach (var group in rendered.GroupBy(item => item.Path.FillGroup)) {
            if (fill) {
                var commands = new List<OfficePathCommand>();
                foreach (var item in group.Where(item => item.Path.CanFill)) AppendPath(commands, item.Points, true);
                if (commands.Count > 0) AddPath(drawing, commands, shape.FillColor, null, 0, OfficeStrokeDashStyle.Solid);
            }
            if (stroke) foreach (var item in group.Where(item => !item.Path.NoLine && item.Points.Count >= 2)) {
                var commands = new List<OfficePathCommand>(); AppendPath(commands, item.Points, item.Path.IsClosed);
                AddPath(drawing, commands, null, shape.LineColor, shape.LineWeight * projection.PhysicalDensity,
                    OfficeStrokeDashStyleMapper.FromVisioLinePattern(shape.LinePattern));
            }
        }
    }

    private static bool HasCurveRows(IEnumerable<XElement> sections) => sections.SelectMany(section => section.Elements())
        .Any(row => ((string?)row.Attribute("T")) is not (null or "MoveTo" or "LineTo" or "RelMoveTo" or "RelLineTo"));

    private static void AppendPath(List<OfficePathCommand> commands, IReadOnlyList<OfficePoint> points, bool close) {
        if (points.Count < 2) return;
        commands.Add(OfficePathCommand.MoveTo(points[0].X, points[0].Y));
        for (int index = 1; index < points.Count; index++) commands.Add(OfficePathCommand.LineTo(points[index].X, points[index].Y));
        if (close) commands.Add(OfficePathCommand.Close());
    }

    private static void AddPath(OfficeDrawing drawing, List<OfficePathCommand> commands, OfficeColor? fill,
        OfficeColor? stroke, double weight, OfficeStrokeDashStyle dash, bool rounded = false) {
        // Cached geometry is flattened to move/line commands. Keep off-page coordinates
        // in the placement transform so each path has valid, non-negative local bounds.
        OfficePoint[] points = commands.Where(command => command.Kind != OfficePathCommandKind.Close)
            .Select(command => command.Point).ToArray();
        bool overflow = points.Any(point => point.X < 0 || point.Y < 0 || point.X > drawing.Width || point.Y > drawing.Height);
        double x = overflow ? points.Min(point => point.X) : 0D;
        double y = overflow ? points.Min(point => point.Y) : 0D;
        double width = overflow ? Math.Max(1D, points.Max(point => point.X) - x) : drawing.Width;
        double height = overflow ? Math.Max(1D, points.Max(point => point.Y) - y) : drawing.Height;
        OfficeShape geometry = OfficeShape.Path(width, height, overflow ? commands.Select(command => command.Translate(x, y)) : commands);
        geometry.FillColor = fill; geometry.StrokeColor = stroke; geometry.StrokeWidth = weight;
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

    private void ReportMetadata(int dataCount, int links, string location) {
        if (dataCount > 0 || links > 0)
            _report.Add("VISIO_DRAWING_METADATA", "Shape data and hyperlink actions are retained in the source but are not included in diagram page content.", OfficeConversionLossKind.Omission, location);
    }
}
