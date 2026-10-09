using OfficeIMO.Drawing;

namespace OfficeIMO.Visio;

internal sealed partial class VisioDrawingProjection {
    private void ProjectConnector(VisioPage page, OfficeDrawing drawing, VisioConnector connector,
        VisioRenderProjection projection, VisioNativeTextStyleResolver styles, string location) {
        List<(double X, double Y)> cached = VisioConnectorGeometry.GetPoints(connector);
        ChargePoints(cached.Count);
        List<OfficePoint> points = cached.Select(point => {
            (double x, double y) = projection.PagePoint(point.X, point.Y); return new OfficePoint(x, y);
        }).ToList();
        double weight = connector.LineWeight * projection.PhysicalDensity;
        bool projectedNative = ProjectConnectorArtwork(drawing, connector, projection, weight, location);
        if (VisioConnectorGeometry.HasVisibleLine(connector) && points.Count >= 2) {
            if (!projectedNative) {
                var commands = new List<OfficePathCommand>(); AppendPath(commands, points, false);
                AddPath(drawing, commands, null, connector.LineColor, weight,
                    OfficeStrokeDashStyleMapper.FromVisioLinePattern(connector.LinePattern), rounded: true);
            }
            ProjectArrow(drawing, points, connector.BeginArrow, true, connector.LineColor, weight, location);
            ProjectArrow(drawing, points, connector.EndArrow, false, connector.LineColor, weight, location);
        }
        if (!string.IsNullOrEmpty(connector.Label)) {
            VisioRenderConnectorLabelPlacement label = VisioRenderLabelLayout.ResolveUnadjusted(connector, cached, projection);
            (double cx, double cy) = VisioConnectorGeometry.GetLabelCenter(connector, label.X, label.Y, label.Width, label.Height);
            (double x, double y) = projection.PagePoint(cx, cy);
            ProjectText(drawing, connector.Label!, connector.TextStyle,
                VisioRichTextProjection.Create(page, connector, 72D, _token, styles), x, y,
                label.Width * projection.GeometryDensity, label.Height * projection.GeometryDensity,
                VisioConnectorLabelFrame.ResolveAngle(connector), 9D, location);
        }
        bool ambiguousNativeRoute = connector.NativeGeometry?.AppliesTo(connector) == true && !connector.NativeGeometry.TryGetPoints(connector, out _);
        _report.Add("VISIO_DRAWING_CONNECTOR_ROUTE", ambiguousNativeRoute
            ? "Native outlines are projected separately; cached endpoints provide the route for labels and arrows where multiple outlines do not identify a unique route."
            : "Cached connector routes and label anchors are projected without ShapeSheet recalculation or collision adjustment.",
            ambiguousNativeRoute ? OfficeConversionLossKind.Approximation : OfficeConversionLossKind.None, location);
        ReportMetadata(connector.ShapeData.Count + connector.Data.Count, connector.Hyperlinks.Count, location);
        if (connector.PreservedShapeChildren.Any(entry => entry.RawElement != null))
            _report.Add("VISIO_DRAWING_CONNECTOR_CONTENT", "Opaque connector children are retained in the source but are not projected.", OfficeConversionLossKind.Omission, location);
    }

    private bool ProjectConnectorArtwork(OfficeDrawing drawing, VisioConnector connector, VisioRenderProjection projection,
        double weight, string location) {
        VisioConnectorNativeGeometry? native = connector.NativeGeometry;
        if (native?.AppliesTo(connector) != true) return false;
        VisioShape shape = native.CreateShape(connector);
        if (!VisioShapeGeometry.TryGetPreservedClosedPaths(shape, out List<VisioShapeGeometryPath> paths)) {
            _report.Add("VISIO_DRAWING_CONNECTOR_GEOMETRY", "Cached endpoints replace native connector geometry that the shared geometry owner cannot project.", OfficeConversionLossKind.Approximation, location);
            return false;
        }
        if (paths.Any(path => path.CanFill))
            _report.Add("VISIO_DRAWING_CONNECTOR_FILL", "Native connector fill artwork is retained in the source but is not projected.", OfficeConversionLossKind.Omission, location);
        if (HasCurveRows(connector.PreservedGeometrySections))
            _report.Add("VISIO_DRAWING_CURVES", "Cached connector curve rows are flattened by the shared Visio geometry owner.", OfficeConversionLossKind.Approximation, location);
        foreach (VisioShapeGeometryPath path in paths.Where(path => !path.NoLine)) {
            ChargePoints(path.Points.Count);
            if (connector.LinePattern == 0 || weight <= 0 || connector.LineColor.A == 0) continue;
            List<OfficePoint> projected = path.Points.Select(point => {
                (double x, double y) = native.ToPage(shape, point.X, point.Y);
                (x, y) = projection.PagePoint(x, y);
                return new OfficePoint(x, y);
            }).ToList();
            var commands = new List<OfficePathCommand>(); AppendPath(commands, projected, path.IsClosed);
            if (commands.Count > 0) AddPath(drawing, commands, null, connector.LineColor, weight,
                OfficeStrokeDashStyleMapper.FromVisioLinePattern(connector.LinePattern), rounded: true);
        }
        return true;
    }

    private void ProjectArrow(OfficeDrawing drawing, List<OfficePoint> points, EndArrow? arrow, bool start,
        OfficeColor color, double weight, string location) {
        if (!arrow.HasValue || arrow.Value == EndArrow.None) return;
        var tuples = points.Select(point => (point.X, point.Y)).ToList();
        if (!OfficeGeometry.TryGetArrowheadSegment(tuples, start, out var tip, out var from) ||
            !OfficeGeometry.TryCreateArrowheadPoints(new OfficePoint(tip.X, tip.Y), new OfficePoint(from.X, from.Y),
                weight, out OfficePoint[] head, minimumLength: 6D)) return;
        ChargePoints(head.Length);
        var commands = new List<OfficePathCommand>(); AppendPath(commands, head, true);
        AddPath(drawing, commands, color, null, 0, OfficeStrokeDashStyle.Solid);
        _report.Add("VISIO_DRAWING_ARROW", "The source arrow kind uses a shared triangular arrowhead approximation.", OfficeConversionLossKind.Approximation, location);
    }
}
