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
                AddPath(drawing, commands, default, connector.LineColor, weight,
                    OfficeStrokeDashStyleMapper.FromVisioLinePattern(connector.LinePattern), rounded: true);
            }
            ProjectArrow(drawing, points, connector.BeginArrow, ArrowSize(connector, "BeginArrowSize"), true, connector.LineColor, weight, location);
            ProjectArrow(drawing, points, connector.EndArrow, ArrowSize(connector, "EndArrowSize"), false, connector.LineColor, weight, location);
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
        // The route's padded bounds are the link target; a thin line alone would be hard to activate.
        double reach = Math.Max(weight * 2D, 3D);
        ProjectLinks(drawing, connector.Hyperlinks, points.SelectMany(point => new[] {
            new OfficePoint(point.X - reach, point.Y - reach), new OfficePoint(point.X + reach, point.Y + reach) }).ToList(), location);
        ReportShapeData(connector.ShapeData.Count + connector.Data.Count, location);
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
        if (paths.Any(path => path.HasFlattenedCurves))
            _report.Add("VISIO_DRAWING_CURVES", "Connector spline and NURBS rows are flattened by the shared Visio geometry owner; arcs and Bezier rows remain exact curves.", OfficeConversionLossKind.Approximation, location);
        OfficePoint Map(double x, double y) {
            (double px, double py) = native.ToPage(shape, x, y);
            (px, py) = projection.PagePoint(px, py);
            return new OfficePoint(px, py);
        }
        foreach (VisioShapeGeometryPath path in paths.Where(path => !path.NoLine)) {
            ChargePoints(path.Points.Count);
            if (connector.LinePattern == 0 || weight <= 0 || connector.LineColor.A == 0) continue;
            AddPath(drawing, CreateCommands(path, Map, path.IsClosed), default, connector.LineColor, weight,
                OfficeStrokeDashStyleMapper.FromVisioLinePattern(connector.LinePattern), rounded: true);
        }
        return true;
    }

    private static double ArrowSize(VisioConnector connector, string cell) {
        string? raw = connector.PreservedCellElements.FirstOrDefault(element => (string?)element.Attribute("N") == cell)?.Attribute("V")?.Value;
        // BeginArrowSize/EndArrowSize: 0 very small ... 2 medium (default) ... 6 colossal.
        return VisioShapeGeometry.TryParseLiteralWithoutShape(raw, out double size) && size >= 0 && size <= 6
            ? new[] { 0.6D, 0.8D, 1D, 1.3D, 1.7D, 2.3D, 3D }[(int)size] : 1D;
    }

    /// <summary>Maps the arrowhead indices whose native shape is known to a shared marker; others return false.</summary>
    internal static bool TryMapArrow(int kind, out OfficeLineMarkerKind marker, out double widthRatio) {
        (marker, widthRatio) = kind switch {
            1 => (OfficeLineMarkerKind.Arrow, 0.9D),
            4 or 5 or 13 => (OfficeLineMarkerKind.Triangle, 0.75D),
            20 => (OfficeLineMarkerKind.Oval, 1D),
            22 => (OfficeLineMarkerKind.Diamond, 0.7D),
            _ => (OfficeLineMarkerKind.None, 0D)
        };
        return marker != OfficeLineMarkerKind.None;
    }

    private void ProjectArrow(OfficeDrawing drawing, List<OfficePoint> points, EndArrow? arrow, double size, bool start,
        OfficeColor color, double weight, string location) {
        if (!arrow.HasValue || arrow.Value == EndArrow.None) return;
        var tuples = points.Select(point => (point.X, point.Y)).ToList();
        if (!OfficeGeometry.TryGetArrowheadSegment(tuples, start, out var tip, out var from)) return;
        int kind = (int)arrow.Value;
        bool mapped = TryMapArrow(kind, out OfficeLineMarkerKind markerKind, out double widthRatio);
        if (!mapped) { markerKind = OfficeLineMarkerKind.Triangle; widthRatio = 0.75D; }
        double length = Math.Max(weight * 4D, 6D) * size;
        var direction = new OfficePoint(tip.X - from.X, tip.Y - from.Y);
        IReadOnlyList<OfficePoint> contour = OfficeLineMarkerGeometry.CreateContour(
            new OfficeLineMarker(markerKind, length * widthRatio, length), new OfficePoint(tip.X, tip.Y), direction);
        if (contour.Count < 3) return;
        ChargePoints(contour.Count);
        var commands = new List<OfficePathCommand>();
        if (markerKind == OfficeLineMarkerKind.Arrow) {
            AppendPath(commands, contour, false);
            AddPath(drawing, commands, default, color, weight, OfficeStrokeDashStyle.Solid, rounded: true);
        } else {
            AppendPath(commands, contour, true);
            AddPath(drawing, commands, new FillPaint(color, null, null, default), null, 0, OfficeStrokeDashStyle.Solid);
        }
        _report.Add(mapped ? "VISIO_DRAWING_ARROW" : "VISIO_DRAWING_ARROW_KIND", mapped
            ? "The arrow kind uses its shared marker shape; native arrow proportions are approximated."
            : $"Arrow kind {kind} has no qualified shape and uses a filled triangle.",
            OfficeConversionLossKind.Approximation, location);
    }
}
